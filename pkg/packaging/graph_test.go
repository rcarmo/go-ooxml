package packaging

import "testing"

func TestGraphURIResolutionBatch(t *testing.T) {
	valid := []struct{ source, target, want string }{{"word/document.xml", "media/a.png", "word/media/a.png"}, {"word/document.xml", "../custom/item.xml", "custom/item.xml"}, {"word/document.xml", "#bookmark", "word/document.xml"}, {"", "/word/document.xml", "word/document.xml"}, {"word/document.xml", "media/a%20b.png", "word/media/a b.png"}}
	for _, tc := range valid {
		t.Run(tc.target, func(t *testing.T) {
			got, err := resolveGraphTarget(tc.source, tc.target)
			if err != nil || got != tc.want {
				t.Fatalf("%q %v", got, err)
			}
		})
	}
	for _, target := range []string{"../../escape", "https://example.com/a", "//host/a", "media/a%2fb.png", `a\b`, "?query=x", "../%2e%2e/x"} {
		t.Run(target, func(t *testing.T) {
			if _, err := resolveGraphTarget("word/document.xml", target); err == nil {
				t.Fatal("unsafe target accepted")
			}
		})
	}
	for _, name := range []string{"weird.rels", "word/rels.xml.rels", "word/_rels/.rels"} {
		t.Run(name, func(t *testing.T) {
			if _, err := sourceForRegistry(name); err == nil {
				t.Fatal("invalid registry path")
			}
		})
	}
}
