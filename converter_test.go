package office_to_pdf

import (
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestConvertXlsxToPdf(t *testing.T) {
	// 1. Create a temporary Excel file for testing
	tmpDir := t.TempDir()
	xlsxPath := filepath.Join(tmpDir, "test.xlsx")
	pdfPath := filepath.Join(tmpDir, "test.pdf")

	f := excelize.NewFile()
	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Test Header")
	f.SetCellValue(sheet, "A2", "Test Data")
	if err := f.SaveAs(xlsxPath); err != nil {
		t.Fatalf("failed to create test xlsx: %v", err)
	}

	// 2. Initialize converter (using embedded font) and run conversion
	converter := NewConverter("", "Arial")
	err := converter.ConvertXlsxToPdf(xlsxPath, pdfPath)
	if err != nil {
		t.Errorf("ConvertXlsxToPdf failed: %v", err)
	}

	// 3. Verify PDF was created
	if _, err := os.Stat(pdfPath); os.IsNotExist(err) {
		t.Error("output PDF file was not created")
	}
}

func TestConvertDocxToPdf(t *testing.T) {
	tmpDir := t.TempDir()
	pdfPath := filepath.Join(tmpDir, "output.pdf")
	
	converter := NewConverter("", "")
	
	// Test with non-existent file to ensure error handling
	err := converter.ConvertDocxToPdf("non-existent.docx", pdfPath)
	if err == nil {
		t.Error("expected error for non-existent docx file, got nil")
	}
}

func TestConvertImageToPdf(t *testing.T) {
	tmpDir := t.TempDir()
	imgPath := filepath.Join(tmpDir, "test.png")

	// 1. Create a dummy image
	createDummyImage(t, imgPath, 100, 100)

	modes := []string{"fit", "stretch", "center", "original"}

	for _, mode := range modes {
		t.Run("Mode_"+mode, func(t *testing.T) {
			pdfPath := filepath.Join(tmpDir, "test_"+mode+".pdf")
			converter := NewConverter("", "")
			converter.ImageFitMode = mode
			err := converter.ConvertImageToPdf(imgPath, pdfPath)
			if err != nil {
				t.Errorf("ConvertImageToPdf failed with mode %s: %v", mode, err)
			}

			if _, err := os.Stat(pdfPath); os.IsNotExist(err) {
				t.Errorf("output PDF file was not created for mode %s", mode)
			}
		})
	}
}

func TestConvertImagesToPdf(t *testing.T) {
	tmpDir := t.TempDir()
	img1 := filepath.Join(tmpDir, "test1.png")
	img2 := filepath.Join(tmpDir, "test2.png")
	pdfPath := filepath.Join(tmpDir, "test.pdf")

	// Create one small and one large image
	createDummyImage(t, img1, 100, 100)
	createDummyImage(t, img2, 1000, 1500) // Exceeds A4

	converter := NewConverter("", "")
	converter.ImageFitMode = "center"
	err := converter.ConvertImagesToPdf([]string{img1, img2}, pdfPath)
	if err != nil {
		t.Errorf("ConvertImagesToPdf failed: %v", err)
	}

	if _, err := os.Stat(pdfPath); os.IsNotExist(err) {
		t.Error("output PDF file was not created")
	}
}

func createDummyImage(t *testing.T, path string, w, h int) {
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			img.Set(x, y, color.RGBA{255, 0, 0, 255})
		}
	}
	f, err := os.Create(path)
	if err != nil {
		t.Fatalf("failed to create test image %s: %v", path, err)
	}
	if err := png.Encode(f, img); err != nil {
		f.Close()
		t.Fatalf("failed to encode test image %s: %v", path, err)
	}
	f.Close()
}
