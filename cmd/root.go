package cmd

import (
	"log/slog"
	"os"

	"github.com/spf13/cobra"
)

var (
	outputDir string
	verbose   bool
	format    string
)

var rootCmd = &cobra.Command{
	Use:   "vba-extract",
	Short: "Extract VBA source from Access MDB/ACCDB files on Linux",
	Long:  "vba-extract extracts and inspects VBA projects embedded in Microsoft Access databases.",
	PersistentPreRun: func(_ *cobra.Command, _ []string) {
		var h slog.Handler
		if verbose {
			h = slog.NewTextHandler(os.Stderr, &slog.HandlerOptions{Level: slog.LevelDebug})
		} else {
			h = slog.DiscardHandler
		}

		slog.SetDefault(slog.New(h))
	},
}

func Execute() error {
	return rootCmd.Execute()
}

func init() {
	rootCmd.PersistentFlags().StringVar(&outputDir, "output-dir", "", "Output directory for extracted files")
	rootCmd.PersistentFlags().BoolVar(&verbose, "verbose", false, "Enable verbose output")
	rootCmd.PersistentFlags().StringVar(&format, "format", "tree", "Output format (flat|tree)")
}
