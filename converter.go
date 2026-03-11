package office_to_pdf

import (
	"embed"
	"fmt"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"os"

	"github.com/ryugenxd/docx2pdf"
	"github.com/signintech/gopdf"
	"github.com/xuri/excelize/v2"
)

//go:embed fonts/arial.ttf
var defaultFont embed.FS

const defaultFontName = "fonts/arial.ttf"

// Converter provides methods to convert Office documents and images to PDF using pure Go.
type Converter struct {
	FontPath     string // Path to a TTF font file
	FontFamily   string // Font family name to use in PDF
	// ImageFitMode determines how images are placed on the PDF page.
	// Supported modes:
	// - "fit" (default): Scaled to fit within the page margins while maintaining aspect ratio, centered.
	// - "stretch": Scaled to fill the entire page within margins, ignoring aspect ratio.
	// - "center": Original size, centered on the page. If larger than the page, it scales down to fit.
	// - "original": Original size, placed at the top-left corner.
	ImageFitMode string
	// Orientation determines the page orientation for PDF conversion (Excel and Images).
	// Supported modes:
	// - "portrait" (default): Standard vertical orientation.
	// - "landscape": Horizontal orientation.
	// - "auto": Determines orientation automatically based on content (Excel columns or Image aspect ratio).
	Orientation  string
}

// NewConverter returns a new Converter with optional font settings.
func NewConverter(fontPath, fontFamily string) *Converter {
	return &Converter{
		FontPath:     fontPath,
		FontFamily:   fontFamily,
		ImageFitMode: "fit",
		Orientation:  "portrait",
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

		totalWidth := 0.0
		for _, w := range colWidths {
			if w < 60 {
				w = 60
			}
			totalWidth += w
		}

		isLandscape := false
		if c.Orientation == "landscape" {
			isLandscape = true
		} else if c.Orientation == "auto" {
			if totalWidth+60 > 595.28 { // 60 for left+right margins
				isLandscape = true
			}
		}

		pageOpt := gopdf.PageOption{PageSize: gopdf.PageSizeA4}
		pageLimitY := 800.0
		if isLandscape {
			pageOpt.PageSize = &gopdf.Rect{W: 841.89, H: 595.28}
			pageLimitY = 550.0 // Adjusted for landscape height
		}

		pdf.AddPageWithOption(pageOpt)
		pdf.SetFont(family, "", 10)

		y := 30.0
		for _, row := range rows {
			x := 30.0
			rowHeight := 20.0
			
			// Check for page break
			if y + rowHeight > pageLimitY {
				pdf.AddPageWithOption(pageOpt)
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

	margin := 20.0

	for _, inputPath := range inputPaths {
		var imgW, imgH float64
		file, err := os.Open(inputPath)
		if err == nil {
			config, _, errConfig := image.DecodeConfig(file)
			file.Close()
			if errConfig == nil {
				// Assuming standard 96 DPI, 1 pixel is approx 0.75 PDF points (72 points per inch)
				imgW = float64(config.Width) * 0.75
				imgH = float64(config.Height) * 0.75
			}
		}

		isLandscape := false
		if c.Orientation == "landscape" {
			isLandscape = true
		} else if c.Orientation == "auto" {
			if imgW > imgH {
				isLandscape = true
			}
		}

		pageOpt := gopdf.PageOption{PageSize: gopdf.PageSizeA4}
		pageW := gopdf.PageSizeA4.W
		pageH := gopdf.PageSizeA4.H
		
		if isLandscape {
			pageW = 841.89
			pageH = 595.28
			pageOpt.PageSize = &gopdf.Rect{W: pageW, H: pageH}
		}
		
		pdf.AddPageWithOption(pageOpt)

		targetW := pageW - (margin * 2)
		targetH := pageH - (margin * 2)

		var rect *gopdf.Rect
		x := margin
		y := margin

		mode := c.ImageFitMode
		if mode == "" {
			mode = "fit"
		}

		if mode == "stretch" || (imgW == 0 && imgH == 0) {
			rect = &gopdf.Rect{W: targetW, H: targetH}
		} else {
			switch mode {
			case "fit":
				ratio := imgW / imgH
				pageRatio := targetW / targetH
				if ratio > pageRatio {
					rect = &gopdf.Rect{W: targetW, H: targetW / ratio}
				} else {
					rect = &gopdf.Rect{W: targetH * ratio, H: targetH}
				}
				x = margin + (targetW-rect.W)/2
				y = margin + (targetH-rect.H)/2
			case "center":
				if imgW > targetW || imgH > targetH {
					// Scale down if it exceeds margins
					ratio := imgW / imgH
					pageRatio := targetW / targetH
					if ratio > pageRatio {
						rect = &gopdf.Rect{W: targetW, H: targetW / ratio}
					} else {
						rect = &gopdf.Rect{W: targetH * ratio, H: targetH}
					}
				} else {
					rect = &gopdf.Rect{W: imgW, H: imgH}
				}
				x = margin + (targetW-rect.W)/2
				y = margin + (targetH-rect.H)/2
			case "original":
				rect = &gopdf.Rect{W: imgW, H: imgH}
			}
		}

		// gopdf.Image automatically scales if a rect is provided.
		err = pdf.Image(inputPath, x, y, rect)
		if err != nil {
			// If scaled image fails, try basic fallback
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
