// Package docxr provides functionality for reading, manipulating, and writing Word documents (DOCX).
package docxr

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"io/fs"
	"os"
	"strings"
	"text/template"

	"github.com/beevik/etree"
)



// Docx represents a Word document.
// It is used to read, manipulate, and write Word documents.
// It is not thread-safe.
type Docx struct {
	templateData map[string]*etree.Document
	files        []string
	src          fs.FS
}

// NewDocx creates and returns a new empty Docx document.
func NewDocx() *Docx {
	return &Docx{
		templateData: make(map[string]*etree.Document),
	}
}

// ReadFromFile reads a DOCX document from the specified file path.
// It returns an error if the file cannot be opened or read.
func (dx *Docx) ReadFromFile(path string) error {
	b, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	return dx.ReadFromBytes(b)
}

// ReadFrom reads a DOCX document from the provided io.Reader.
// It returns the number of bytes read and an error, if any.
// Note: For non-seekable readers, the entire content is read into memory,
// which may be inefficient for very large files.
func (dx *Docx) ReadFrom(r io.Reader) (int64, error) {
	ra, size, err := toReaderAt(r)
	if err != nil {
		return 0, err
	}
	return size, dx.readFrom(ra, size)
}

// ReadFromBytes reads a DOCX document from the provided byte slice.
// It returns an error if the byte slice cannot be parsed as a DOCX.
func (dx *Docx) ReadFromBytes(b []byte) error {
	ra := bytes.NewReader(b)
	return dx.readFrom(ra, int64(len(b)))
}

func (dx *Docx) readFrom(ra io.ReaderAt, size int64) error {
	zipReader, err := zip.NewReader(ra, size)
	if err != nil {
		return err
	}
	return dx.unpack(zipReader)
}

func (dx *Docx) unpack(zr *zip.Reader) error {
	for _, f := range zr.File {
		if _, exist := supportedFiles[f.Name]; exist {
			doc := etree.NewDocument()
			zf, err := f.Open()
			if err != nil {
				return err
			}
			defer zf.Close()
			_, err = doc.ReadFrom(zf)
			if err != nil {
				return err
			}
			dx.templateData[f.Name] = doc
		}
		dx.files = append(dx.files, f.Name)
	}

	dx.src = zr

	return nil
}

// Render applies the provided data to the document's templates.
// It processes all supported XML files within the DOCX (e.g., document.xml, headers, footers)
// and replaces template tags (e.g., `{{.FieldName}}`) with corresponding values from `data`.
// The `data` parameter can be any type that is compatible with Go's text/template package.
// It returns an error if templating fails.
func (dx *Docx) Render(data any) error {
	newRenderMap := make(map[string]*etree.Document, len(dx.templateData))

	for key, et := range dx.templateData {
		mergeTemplateTags(et)
		newDoc, err := dx.render(et, key, data)
		if err != nil {
			return err
		}
		newRenderMap[key] = newDoc
	}

	dx.templateData = newRenderMap
	return nil
}

func (dx *Docx) render(et *etree.Document, name string, data any) (*etree.Document, error) {
	var (
		strBuf        = new(bytes.Buffer)
		newContentBuf = new(bytes.Buffer)
	)
	_, err := et.WriteTo(strBuf)
	if err != nil {
		return nil, fmt.Errorf("docxr: failed to write xml to buffer for template parsing: %w", err)
	}
	tmpl, err := template.New(name).Parse(strBuf.String())
	if err != nil {
		return nil, fmt.Errorf("docxr: failed to parse template: %w", err)
	}

	err = tmpl.Execute(newContentBuf, data)
	if err != nil {
		return nil, fmt.Errorf("docxr: failed to execute template: %w", err)
	}
	newDoc := etree.NewDocument()
	_, err = newDoc.ReadFrom(newContentBuf)
	if err != nil {
		return nil, fmt.Errorf("docxr: failed to read rendered content: %w", err)
	}

	return newDoc, nil
}

func (dx *Docx) pack(zw *zip.Writer) error {
	for _, fileName := range dx.files {
		w, err := zw.Create(fileName)
		if err != nil {
			return err
		}

		// If the file was parsed for templating, write the (potentially modified) content.
		if doc, ok := dx.templateData[fileName]; ok {
			if _, err := doc.WriteTo(w); err != nil {
				return fmt.Errorf("docxr: failed to write rendered xml to zip for %s: %w", fileName, err)
			}
			continue
		}

		// For directories, creating the entry is sufficient.
		if strings.HasSuffix(fileName, "/") {
			continue
		}

		// Otherwise, copy the original file content from the source.
		// Use an anonymous function to leverage defer for proper resource cleanup.
		if err := func() error {
			srcFile, err := dx.src.Open(fileName)
			if err != nil {
				return fmt.Errorf("docxr: failed to open original file %s from zip: %w", fileName, err)
			}
			defer srcFile.Close()

			if _, err := io.Copy(w, srcFile); err != nil {
				return fmt.Errorf("docxr: failed to copy original file %s to zip: %w", fileName, err)
			}
			return nil
		}(); err != nil {
			return err
		}
	}
	return nil
}

// WriteTo writes the current state of the DOCX document to the provided io.Writer.
// This includes both the templated content and original files.
// It returns the number of bytes written and an error, if any.
func (dx *Docx) WriteTo(w io.Writer) (_ int64, err error) {
	zipWriter := zip.NewWriter(w)
	defer zipWriter.Close()

	return 0, dx.pack(zipWriter)
}

// WriteToFile writes the current state of the DOCX document to the specified file path.
// It returns an error if the file cannot be created or written to.
func (dx *Docx) WriteToFile(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	_, err = dx.WriteTo(f)
	return err
}

// toReaderAt converts an io.Reader to an io.ReaderAt.
// It first attempts a direct type assertion for efficiency. If the reader
// does not already support seeking, it falls back to reading the entire
// content into memory, which may be inefficient for large files.
func toReaderAt(r io.Reader) (io.ReaderAt, int64, error) {
	// If r already implements io.ReaderAt and io.Seeker, we can use it directly.
	// This is a fast path for types like *os.File, which are passed from ReadFromFile.
	if ra, ok := r.(io.ReaderAt); ok {
		if seeker, ok := r.(io.Seeker); ok {
			size, err := seeker.Seek(0, io.SeekEnd)
			if err != nil {
				return nil, 0, err
			}
			_, err = seeker.Seek(0, io.SeekStart)
			if err != nil {
				return nil, 0, err
			}
			return ra, size, nil
		}
	}

	// Fallback for readers that don't support seeking (e.g. network streams).
	data, err := io.ReadAll(r)
	if err != nil {
		return nil, 0, err
	}
	return bytes.NewReader(data), int64(len(data)), nil
}
