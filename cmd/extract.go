package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

var extractCmd = &cobra.Command{
	Use:   "extract [files...]",
	Short: "Extract VBA modules from one or more Access files",
	Args:  cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		fmt.Printf("extract not implemented yet; files=%v output-dir=%q format=%q verbose=%v\n", args, outputDir, format, verbose)
		return nil
	},
}

func init() {
	rootCmd.AddCommand(extractCmd)
}
