package main

import (
	"fmt"
	"os"

	"github.com/MeKo-Christian/accessdump/internal/mdb"
)

func main() {
	db, err := mdb.Open("testdata/Start.mdb")
	if err != nil { fmt.Fprintf(os.Stderr, "open: %v\n", err); os.Exit(1) }
	defer db.Close()

	images, err := mdb.ExtractImages(db)
	if err != nil { fmt.Fprintf(os.Stderr, "extract: %v\n", err); os.Exit(1) }

	for i, img := range images {
		fmt.Printf("[%d] FormName=%q FileName=%q Format=%s Size=%d\n", i, img.FormName, img.FileName, img.Format, len(img.Data))
	}
}
