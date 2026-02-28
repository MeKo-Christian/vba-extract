package vba

import (
	"bytes"
	"fmt"
	"log/slog"
	"sort"
	"strings"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
)

// StorageNode is one entry from MSysAccessStorage.
type StorageNode struct {
	ID         int32
	ParentID   int32
	Name       string
	Type       int32
	DateCreate any
	DateUpdate any
	LvRaw      []byte
	Data       []byte
	ResolveErr error
}

// StorageTree contains all MSysAccessStorage rows with helper indexes.
type StorageTree struct {
	Nodes    []*StorageNode
	ByID     map[int32]*StorageNode
	Children map[int32][]*StorageNode
}

// LoadStorageTree reads MSysAccessStorage and builds an in-memory tree.
// It also resolves Lv references to raw stream data in node.Data.
func LoadStorageTree(db *mdb.Database) (*StorageTree, error) {
	td, err := findMSysAccessStorageTable(db)
	if err != nil {
		return nil, err
	}

	rows, err := td.ReadRows()
	if err != nil {
		return nil, fmt.Errorf("vba: read MSysAccessStorage rows: %w", err)
	}

	st := &StorageTree{
		ByID:     make(map[int32]*StorageNode),
		Children: make(map[int32][]*StorageNode),
	}

	for _, row := range rows {
		node, err := rowToNode(db, row)
		if err != nil {
			return nil, err
		}
		if node == nil {
			continue
		}

		if node.ParentID == node.ID {
			node.ParentID = 0
		}

		st.Nodes = append(st.Nodes, node)
		st.ByID[node.ID] = node
		st.Children[node.ParentID] = append(st.Children[node.ParentID], node)
	}

	for parentID := range st.Children {
		sort.Slice(st.Children[parentID], func(i, j int) bool {
			if st.Children[parentID][i].ID == st.Children[parentID][j].ID {
				return st.Children[parentID][i].Name < st.Children[parentID][j].Name
			}
			return st.Children[parentID][i].ID < st.Children[parentID][j].ID
		})
	}

	return st, nil
}

func findMSysAccessStorageTable(db *mdb.Database) (*mdb.TableDef, error) {
	if td, err := db.FindTable("MSysAccessStorage"); err == nil {
		return td, nil
	}

	entries, err := db.Catalog()
	if err != nil {
		return nil, fmt.Errorf("vba: locate MSysAccessStorage: %w", err)
	}

	for _, entry := range entries {
		if entry.Type == mdb.ObjTypeLocalTable && strings.EqualFold(entry.Name, "MSysAccessStorage") {
			return db.ReadTableDef(int64(entry.ID))
		}
	}

	return nil, fmt.Errorf("vba: table MSysAccessStorage not found")
}

func rowToNode(db *mdb.Database, row mdb.Row) (*StorageNode, error) {
	id, ok := asInt32(row["Id"])
	if !ok {
		return nil, nil
	}

	node := &StorageNode{
		ID:         id,
		ParentID:   asInt32Default(row["ParentId"]),
		Name:       asString(row["Name"]),
		Type:       asInt32Default(row["Type"]),
		DateCreate: row["DateCreate"],
		DateUpdate: row["DateUpdate"],
	}

	lvRaw, ok := asBytes(row["Lv"])
	if !ok || len(lvRaw) == 0 {
		return node, nil
	}

	node.LvRaw = lvRaw
	resolved, err := db.ResolveMemo(lvRaw)
	if err != nil {
		node.ResolveErr = fmt.Errorf("vba: resolve Lv for node %d (%q): %w", node.ID, node.Name, err)
		return node, nil
	}
	node.Data = resolved

	return node, nil
}

func asInt32(v interface{}) (int32, bool) {
	switch t := v.(type) {
	case int32:
		return t, true
	case int16:
		return int32(t), true
	case int:
		return int32(t), true
	case uint16:
		return int32(t), true
	case uint32:
		return int32(t), true
	case int64:
		return int32(t), true
	case uint64:
		return int32(t), true
	default:
		return 0, false
	}
}

func asInt32Default(v interface{}) int32 {
	if x, ok := asInt32(v); ok {
		return x
	}
	return 0
}

func asString(v interface{}) string {
	s, _ := v.(string)
	return s
}

func asBytes(v interface{}) ([]byte, bool) {
	b, ok := v.([]byte)
	if !ok {
		return nil, false
	}
	dup := make([]byte, len(b))
	copy(dup, b)
	return dup, true
}

// Root returns the ROOT storage node if present.
func (st *StorageTree) Root() *StorageNode {
	for _, node := range st.Nodes {
		if node.ParentID == 0 {
			return node
		}
	}

	for _, node := range st.Nodes {
		if node.Name == "ROOT" {
			return node
		}
	}

	if len(st.Nodes) == 0 {
		return nil
	}

	// Fallback: choose node with smallest ID.
	root := st.Nodes[0]
	for _, node := range st.Nodes[1:] {
		if node.ID < root.ID {
			root = node
		}
	}

	return root
}

// FindChild returns a direct child with the given name under parentID.
func (st *StorageTree) FindChild(parentID int32, name string) *StorageNode {
	for _, node := range st.Children[parentID] {
		if node.Name == name {
			return node
		}
	}
	return nil
}

// FindChildFold returns a direct child with case-insensitive name matching.
func (st *StorageTree) FindChildFold(parentID int32, name string) *StorageNode {
	for _, node := range st.Children[parentID] {
		if strings.EqualFold(node.Name, name) {
			return node
		}
	}
	return nil
}

// VBAProjectNode returns the VBAProject node under ROOT.
func (st *StorageTree) VBAProjectNode() (*StorageNode, error) {
	root := st.Root()
	if root == nil {
		return nil, fmt.Errorf("vba: ROOT node not found in storage tree")
	}

	vbaProject := st.FindChild(root.ID, "VBAProject")
	if vbaProject == nil {
		vbaProject = st.FindChildFold(root.ID, "VBAProject")
	}
	if vbaProject == nil {
		// Access variants may use ROOT -> VBA -> VBAProject.
		for _, child := range st.Children[root.ID] {
			if strings.EqualFold(child.Name, "VBA") {
				vbaProject = st.FindChildFold(child.ID, "VBAProject")
				if vbaProject != nil {
					break
				}
			}
		}
	}

	if vbaProject == nil {
		for _, node := range st.Nodes {
			if strings.EqualFold(node.Name, "VBAProject") {
				vbaProject = node
				break
			}
		}
	}

	if vbaProject == nil {
		return nil, fmt.Errorf("vba: VBAProject node not found")
	}

	return vbaProject, nil
}

// VBAFolderNode returns the VBA folder under VBAProject.
func (st *StorageTree) VBAFolderNode() (*StorageNode, error) {
	vbaProject, err := st.VBAProjectNode()
	if err != nil {
		return nil, err
	}

	vbaFolder := st.FindChild(vbaProject.ID, "VBA")
	if vbaFolder == nil {
		vbaFolder = st.FindChildFold(vbaProject.ID, "VBA")
	}
	if vbaFolder == nil {
		return nil, fmt.Errorf("vba: VBA folder not found under VBAProject")
	}

	return vbaFolder, nil
}

// RequiredStreams returns required stream nodes under VBA folder.
// Keys: PROJECT, PROJECTwm, dir, _VBA_PROJECT.
func (st *StorageTree) RequiredStreams() (map[string]*StorageNode, error) {
	requiredNames := []string{"PROJECT", "PROJECTwm", "dir", "_VBA_PROJECT"}
	result := make(map[string]*StorageNode, len(requiredNames))

	if vbaProject, err := st.VBAProjectNode(); err == nil {
		subtree := st.subtree(vbaProject.ID)
		for _, name := range requiredNames {
			for _, node := range subtree {
				if strings.EqualFold(node.Name, name) {
					result[name] = node
					break
				}
			}
		}
	}

	for _, name := range requiredNames {
		if result[name] != nil {
			continue
		}
		if node := st.findByNameGlobal(name); node != nil {
			result[name] = node
		}
	}

	if result["PROJECT"] == nil {
		if node := st.findLikelyProjectNode(); node != nil {
			result["PROJECT"] = node
		}
	}
	if result["dir"] == nil {
		if node := st.findLikelyDirNode(); node != nil {
			result["dir"] = node
		}
	}

	if result["PROJECT"] == nil {
		return result, fmt.Errorf("vba: PROJECT stream not found")
	}

	return result, nil
}

func (st *StorageTree) subtree(parentID int32) []*StorageNode {
	var out []*StorageNode
	visited := map[int32]bool{}
	var walk func(int32)

	walk = func(pid int32) {
		if visited[pid] {
			return
		}
		visited[pid] = true

		children := st.Children[pid]
		for _, child := range children {
			if child.ID == pid {
				continue
			}
			out = append(out, child)
			walk(child.ID)
		}
	}

	walk(parentID)
	return out
}

// ModuleStreams returns all child streams under VBA folder excluding required streams.
func (st *StorageTree) ModuleStreams() ([]*StorageNode, error) {
	vbaFolder, err := st.VBAFolderNode()
	if err == nil {
		required := map[string]struct{}{
			"PROJECT":      {},
			"PROJECTwm":    {},
			"dir":          {},
			"_VBA_PROJECT": {},
		}

		var modules []*StorageNode
		for _, node := range st.Children[vbaFolder.ID] {
			if _, isRequired := required[node.Name]; isRequired {
				continue
			}
			modules = append(modules, node)
		}

		sort.Slice(modules, func(i, j int) bool {
			if modules[i].ID == modules[j].ID {
				return modules[i].Name < modules[j].Name
			}
			return modules[i].ID < modules[j].ID
		})

		if len(modules) > 0 {
			return modules, nil
		}
	}

	required, _ := st.RequiredStreams()
	reservedNames := map[string]struct{}{}
	for name := range required {
		reservedNames[strings.ToLower(name)] = struct{}{}
	}

	var modules []*StorageNode
	for _, node := range st.Nodes {
		if node == nil || len(node.Data) == 0 {
			continue
		}
		if _, isReserved := reservedNames[strings.ToLower(node.Name)]; isReserved {
			continue
		}
		if node.Type == 2 || isLikelyModuleStreamName(node.Name) {
			modules = append(modules, node)
		}
	}

	sort.Slice(modules, func(i, j int) bool {
		if modules[i].ID == modules[j].ID {
			return modules[i].Name < modules[j].Name
		}
		return modules[i].ID < modules[j].ID
	})

	if len(modules) == 0 {
		return nil, fmt.Errorf("vba: no module streams found")
	}

	return modules, nil
}

func (st *StorageTree) findByNameGlobal(name string) *StorageNode {
	var best *StorageNode
	for _, node := range st.Nodes {
		if !strings.EqualFold(node.Name, name) {
			continue
		}
		if best == nil {
			best = node
			continue
		}
		if len(node.Data) > len(best.Data) {
			best = node
		}
	}
	return best
}

func (st *StorageTree) findLikelyProjectNode() *StorageNode {
	for _, node := range st.Nodes {
		if len(node.Data) == 0 {
			continue
		}
		data := node.Data
		if bytes.Contains(data, []byte("Module=")) || bytes.Contains(data, []byte("DocClass=")) || bytes.Contains(data, []byte("Class=")) {
			if _, err := ParseProjectStream(data); err == nil {
				return node
			}
		}
	}
	return nil
}

func (st *StorageTree) findLikelyDirNode() *StorageNode {
	for _, node := range st.Nodes {
		if len(node.Data) == 0 {
			continue
		}
		if _, err := ParseDirStream(node.Data, func(in []byte) ([]byte, error) {
			out, _, derr := DecompressContainerWithFallback(in, slog.New(slog.DiscardHandler))
			return out, derr
		}); err == nil {
			return node
		}
	}
	return nil
}

func isLikelyModuleStreamName(name string) bool {
	if name == "" {
		return false
	}
	upper := strings.ToUpper(name)
	if strings.Contains(upper, "VBA") || strings.Contains(upper, "PROJECT") {
		return false
	}
	if len(upper) < 12 {
		return false
	}
	for _, r := range upper {
		if (r < 'A' || r > 'Z') && r != '_' {
			return false
		}
	}
	return true
}
