// Command fixturepath resolves trusted fixture IDs through the pinned shared manifest.
package main

import (
	"fmt"
	"os"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintln(os.Stderr, "usage: fixturepath fixture-<sha256> [...]")
		os.Exit(2)
	}
	for _, id := range os.Args[1:] {
		path, err := testutil.LookupFixture(id)
		if err != nil {
			fmt.Fprintln(os.Stderr, err)
			os.Exit(1)
		}
		fmt.Println(path)
	}
}
