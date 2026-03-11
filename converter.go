package office_to_pdf

import (
	"embed"
	"fmt"

	"github.com/ryugenxd/docx2pdf"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

//go:embed fonts/arial.ttf
var defaultFont embed.FS

const defaultFontName = "fonts/arial.ttf"

// Converter provides methods to convert Office documents and images to PDF using pure Go.
type Converter struct {
	FontPath   string // Path to a TTF font file
	FontFamily string // Font family name to use in PDF
}

// NewConverter returns a new Converter with optional font settings.
func NewConverter(fontPath, fontFamily string) *Converter {
	return &Converter{
		FontPath:   fontPath,
		FontFamily: fontFamily,
	}
}

// ConvertDocxToPdf converts a DOCX file to a PDF file.
func (c *Converter) ConvertDocxToPdf(inputPath, outputPath string) error {
	// Using ryugenxd/docx2pdf as it is a pure Go DOCX to PDF converter.
	err := docx2pdf.ConvertFile(inputPath, outputPath)
	if err != nil {
		return fmt.Errorf("failed to convert docx to pdf: %w", err)
	}
	return nil
}

// ConvertXlsxToPdf converts an XLSX file to a PDF file.
func (c *Converter) ConvertXlsxToPdf(inputPath, outputPath string) error {
	f, err := excelize.OpenFile(inputPath)
	if err != nil {
		return fmt.Errorf("failed to open excel file: %w", err)
	}
	defer f.Close()

	pdf := &gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4})

	family := c.FontFamily
	if family == "" {
		family = "Arial"
	}

	if c.FontPath != "" {
		err = pdf.AddTTFFont(family, c.FontPath)
	} else {
		// Use embedded font as default
		fontBytes, _ := defaultFont.ReadFile(defaultFontName)
		err = pdf.AddTTFFontData(family, fontBytes)
	}

	if err != nil {
		return fmt.Errorf("failed to add font: %w", err)
	}

	for _, sheetName := range f.GetSheetList() {
		pdf.AddPage()
		pdf.SetFont(family, "", 10)

		rows, err := f.GetRows(sheetName)
		if err != nil {
			continue
		}

		// Calculate basic column widths based on content
		colWidths := make(map[int]float64)
		for _, row := range rows {
			for i, cell := range row {
				w := float64(len(cell)) * 7.0 // estimation
				if w > colWidths[i] {
					colWidths[i] = w
				}
				if colWidths[i] > 300 {
					colWidths[i] = 300
				}
			}
		}

		y := 30.0
		for _, row := range rows {
			x := 30.0
			rowHeight := 20.0
			
			// Check for page break
			if y + rowHeight > 800 {
				pdf.AddPage()
				y = 30.0
			}

			for i, cell := range row {
				width := colWidths[i]
				if width < 60 {
					width = 60
				}

				// Draw cell border
				pdf.SetLineWidth(0.5)
				pdf.RectFromUpperLeftWithStyle(x, y, width, rowHeight, "D")
				
				// Draw text with clipping to avoid overflow
				rect := &gopdf.Rect{
					W: width - 4,
					H: rowHeight - 4,
				}
				pdf.SetXY(x+2, y+2)
				pdf.Cell(rect, cell)
				
				x += width
			}
			y += rowHeight
		}
	}

	err = pdf.WritePdf(outputPath)
	if err != nil {
		return fmt.Errorf("failed to write pdf: %w", err)
	}

	return nil
}

// ConvertImageToPdf converts a single image file (JPG, PNG) to a PDF file.
func (c *Converter) ConvertImageToPdf(inputPath, outputPath string) error {
	return c.ConvertImagesToPdf([]string{inputPath}, outputPath)
}

// ConvertImagesToPdf converts multiple image files (JPG, PNG) to a single PDF file (one image per page).
func (c *Converter) ConvertImagesToPdf(inputPaths []string, outputPath string) error {
	pdf := &gopdf.GoPdf{}
	pdf.Start(gopdf.Config{PageSize: *gopdf.PageSizeA4})

	pageW := gopdf.PageSizeA4.W
	pageH := gopdf.PageSizeA4.H
	margin := 20.0
	targetW := pageW - (margin * 2)
	targetH := pageH - (margin * 2)

	for _, inputPath := range inputPaths {
		pdf.AddPage()

		// gopdf.Image automatically scales if a rect is provided.
		err := pdf.Image(inputPath, margin, margin, &gopdf.Rect{W: targetW, H: targetH})
		if err != nil {
			// If scaling fails, try natural size
			err = pdf.Image(inputPath, margin, margin, nil)
			if err != nil {
				return fmt.Errorf("failed to add image %s to pdf: %w", inputPath, err)
			}
		}
	}

	err := pdf.WritePdf(outputPath)
	if err != nil {
		return fmt.Errorf("failed to write pdf: %w", err)
	}

	return nil
}
