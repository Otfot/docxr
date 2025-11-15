// Package docxr provides functionality for reading, manipulating, and writing Word documents (DOCX).
package docxr

import (
	"regexp"

	"github.com/beevik/etree"
)

var tagRegex = regexp.MustCompile("{{.*?}}")
var incompleteTagRegex = regexp.MustCompile("{{[^}]*?(}$|$)|{$|^}")

// mergeTemplateTags merges template tags that have been split across multiple XML elements.
// Word processors can break a single tag like `{{.Name}}` into multiple runs (`<w:t>`),
// for example: `<w:t>{{.Na</w:t><w:t>me}}</w:t>`.
// This function iterates through all `<w:t>` elements and merges these split tags
// back into a single element.
func mergeTemplateTags(et *etree.Document) {
	// Find all w:t elements in the document.
	textElements := et.FindElementsSeq("//w:t")
	if textElements == nil {
		return
	}

	var (
		// incompleteTag is a flag that indicates if we are in the process of merging a tag.
		incompleteTag bool
		// currentText holds the merged text content of a tag being processed.
		currentText string
	)

	for textElement := range textElements {
		text := textElement.Text()

		if incompleteTag { // If we are already inside a split tag.
			currentText += text

			// Check if the merged text still looks like an incomplete tag.
			if hasIncompleteTemplateTag(currentText) {
				// If it's still incomplete, clear the current XML element's text
				// and continue to the next one, accumulating its content.
				textElement.SetText("")
				continue
			}

			// If the merged text now forms a complete tag (e.g., "{{.Name}}").
			if hasTemplateTag(currentText) {
				// Set the complete tag content to the current element.
				textElement.SetText(currentText)
				// Reset the state machine.
				incompleteTag = false
				currentText = ""
			} else {
				// This case should ideally not be reached if the logic is sound,
				// but as a fallback, we clear the text to avoid orphaned content.
				textElement.SetText("")
			}
		} else { // If we are not currently inside a split tag.
			currentText = text
			// Check if the current element's text is an incomplete tag.
			if hasIncompleteTemplateTag(currentText) {
				// If so, start the state machine.
				incompleteTag = true
				// Clear the current element's text, as its content is now in currentText
				// and will be moved to a later element.
				textElement.SetText("")
			}
		}
	}
}

func hasTemplateTag(text string) bool {
	return tagRegex.MatchString(text)
}

func hasIncompleteTemplateTag(text string) bool {
	return incompleteTagRegex.MatchString(text)
}
