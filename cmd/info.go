package cmd

import (
	"fmt"
	"sort"
	"strings"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
	"github.com/MeKo-Tech/vba-extract/internal/vba"
	"github.com/spf13/cobra"
)

var (
	infoShowTree bool
	infoForensic bool
)

var infoCmd = &cobra.Command{
	Use:   "info [file]",
	Short: "Show MDB/ACCDB metadata and VBA project info",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		db, err := mdb.Open(args[0])
		if err != nil {
			return err
		}
		defer db.Close()

		names, _ := db.TableNames()
		sort.Strings(names)

		fmt.Printf("file: %s\n", args[0])
		fmt.Printf("jetVersion: %d\n", db.Header.JetVersion)
		fmt.Printf("codepage: %d\n", db.Header.CodePage)
		fmt.Printf("pageCount: %d\n", db.PageCount())
		fmt.Printf("tableCount: %d\n", len(names))

		st, err := vba.LoadStorageTree(db)
		if err != nil {
			fmt.Printf("vba: unavailable (%v)\n", err)
			return nil
		}

		required, reqErr := st.RequiredStreams()
		if reqErr != nil {
			fmt.Printf("vba: storage present, but required streams unresolved (%v)\n", reqErr)
		} else {
			if projectNode := required["PROJECT"]; projectNode != nil && len(projectNode.Data) > 0 {
				project, parseErr := vba.ParseProjectStream(projectNode.Data)
				if parseErr == nil {
					fmt.Printf("vbaProject: %s\n", project.Name)
					fmt.Printf("vbaModules: %d\n", len(project.Modules))
				} else {
					fmt.Printf("vbaProject: parse error (%v)\n", parseErr)
				}
			}
		}

		if infoShowTree {
			printStorageTree(st)
		}

		if infoForensic {
			printForensic(st)
		}

		return nil
	},
}

func init() {
	rootCmd.AddCommand(infoCmd)
	infoCmd.Flags().BoolVar(&infoShowTree, "tree", false, "Show storage tree")
	infoCmd.Flags().BoolVar(&infoForensic, "forensic", false, "Scan all storage streams for hidden/non-standard VBA candidates")
}

func printStorageTree(st *vba.StorageTree) {
	root := st.Root()
	if root == nil {
		fmt.Println("storageTree: <no root>")
		return
	}

	fmt.Println("storageTree:")

	visited := map[int32]bool{}
	var walk func(node *vba.StorageNode, depth int)
	walk = func(node *vba.StorageNode, depth int) {
		if node == nil {
			return
		}

		if depth > 20 {
			indent := ""

			var indentSb92 strings.Builder
			for range depth {
				indentSb92.WriteString("  ")
			}

			indent += indentSb92.String()

			fmt.Printf("%s- ... (depth limit reached)\n", indent)

			return
		}

		if visited[node.ID] {
			indent := ""

			var indentSb100 strings.Builder
			for range depth {
				indentSb100.WriteString("  ")
			}

			indent += indentSb100.String()

			fmt.Printf("%s- %s (id=%d) [cycle]\n", indent, node.Name, node.ID)

			return
		}

		visited[node.ID] = true

		indent := ""

		var indentSb109 strings.Builder
		for range depth {
			indentSb109.WriteString("  ")
		}

		indent += indentSb109.String()

		fmt.Printf("%s- %s (id=%d type=%d data=%d)\n", indent, node.Name, node.ID, node.Type, len(node.Data))

		for _, child := range st.Children[node.ID] {
			if child != nil && child.ID == node.ID {
				continue
			}

			walk(child, depth+1)
		}

		visited[node.ID] = false
	}

	walk(root, 0)
}

func printForensic(st *vba.StorageTree) {
	report := vba.ForensicScanStorage(st)

	fmt.Println("forensic:")
	fmt.Printf("  hits=%d projectCandidates=%d dirCandidates=%d sourceCandidates=%d compressedCandidates=%d artifactCandidates=%d\n",
		len(report.Hits), report.ProjectCandidates, report.DirCandidates, report.SourceCandidates, report.CompressedCandidates, report.ArtifactCandidates)

	limit := min(len(report.Hits), 25)

	for i := range limit {
		h := report.Hits[i]
		fmt.Printf("  - id=%d name=%q type=%d size=%d kind=%s score=%d :: %s\n",
			h.NodeID, h.NodeName, h.NodeType, h.DataSize, h.Kind, h.Score, h.Summary)
	}

	if len(report.Hits) > limit {
		fmt.Printf("  ... %d more hit(s)\n", len(report.Hits)-limit)
	}
}
