package cmd

import (
	"fmt"
	"sort"

	"github.com/spf13/cobra"
)

var listJSON bool

var listCmd = &cobra.Command{
	Use:   "list [file]",
	Short: "List VBA modules in an Access file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		modules, err := loadModules(args[0])
		if err != nil {
			return err
		}

		entries := make([]listEntry, 0, len(modules))
		for _, m := range modules {
			entries = append(entries, listEntry{
				Name:      m.Name,
				Type:      string(m.Type),
				Stream:    m.Stream,
				SizeBytes: len(m.Text),
				Partial:   m.Partial,
			})
		}

		sort.Slice(entries, func(i, j int) bool {
			if entries[i].Name == entries[j].Name {
				return entries[i].Stream < entries[j].Stream
			}

			return entries[i].Name < entries[j].Name
		})

		if listJSON {
			return printListJSON(entries)
		}

		fmt.Printf("file: %s modules: %d\n", args[0], len(entries))
		printListTable(entries)

		return nil
	},
}

func init() {
	rootCmd.AddCommand(listCmd)
	listCmd.Flags().BoolVar(&listJSON, "json", false, "Output as JSON")
}
