package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var infoCmd = &cobra.Command{
	Use:   "info [file]",
	Short: "Show MDB/ACCDB metadata and VBA project info",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		fmt.Printf("info not implemented yet; file=%q verbose=%v\n", args[0], verbose)
		return nil
	},
}

func init() {
	rootCmd.AddCommand(infoCmd)
}
