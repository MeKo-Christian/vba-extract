package cmd

import (
	"fmt"
	"sort"

	"github.com/MeKo-Tech/vba-extract/internal/mdb"
	"github.com/MeKo-Tech/vba-extract/internal/vba"
	"github.com/spf13/cobra"
)

var infoShowTree bool

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

		return nil
	},
}

func init() {
	rootCmd.AddCommand(infoCmd)
	infoCmd.Flags().BoolVar(&infoShowTree, "tree", false, "Show storage tree")
}

func printStorageTree(st *vba.StorageTree) {
	root := st.Root()
	if root == nil {
		fmt.Println("storageTree: <no root>")
		return
	}

	fmt.Println("storageTree:")
	var walk func(node *vba.StorageNode, depth int)
	walk = func(node *vba.StorageNode, depth int) {
		indent := ""
		for i := 0; i < depth; i++ {
			indent += "  "
		}
		fmt.Printf("%s- %s (id=%d type=%d data=%d)\n", indent, node.Name, node.ID, node.Type, len(node.Data))
		for _, child := range st.Children[node.ID] {
			walk(child, depth+1)
		}
	}

	walk(root, 0)
}
