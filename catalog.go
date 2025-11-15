// Package docxr provides functionality for reading, manipulating, and writing Word documents (DOCX).
package docxr

// These constants define the paths to various XML files within a DOCX archive
// that are supported for templating and rendering by the docxr package.
const (
	WordDocument = "word/document.xml"
	WordHeader1  = "word/header1.xml"
	WordHeader2  = "word/header2.xml"
	WordHeader3  = "word/header3.xml"
	WordFooter1  = "word/footer1.xml"
	WordFooter2  = "word/footer2.xml"
)

// supportedFiles is a map of supported files for rendering
var supportedFiles = map[string]bool{
	WordDocument: true,
	WordHeader1:  true,
	WordHeader2:  true,
	WordHeader3:  true,
	WordFooter1:  true,
	WordFooter2:  true,
}
