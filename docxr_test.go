package docxr

import "testing"

func TestDocxr(t *testing.T) {
	type Hello struct {
		Header string
		World  string
		IsTrue bool
		Items  []string
		Footer string
	}

	hello := Hello{
		Header: "Header",
		World:  "World",
		IsTrue: true,
		Items:  []string{"Item1", "Item2"},
		Footer: "Footer",
	}

	docx := NewDocx()
	docx.ReadFromFile("example/Hello.docx")
	err := docx.Render(hello)
	if err != nil {
		t.Errorf("Error rendering document: %v", err)
	}
	err = docx.WriteToFile("example/Output.docx")
	if err != nil {
		t.Errorf("Error writing document: %v", err)
	}
}
