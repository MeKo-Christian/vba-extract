package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var listCmd = &cobra.Command{
	Use:   "list [file]",
	Short: "List VBA modules in an Access file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		fmt.Printf("list not implemented yet; file=%q verbose=%v\n", args[0], verbose)
		return nil
	},
}

func init() {
	rootCmd.AddCommand(listCmd)
}
