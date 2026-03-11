# Office to PDF Go Module

A reusable Go module for converting Excel (`.xlsx`), Word (`.docx`), and Images (`.jpg`, `.png`) to PDF using pure Go (MIT-licensed) libraries.

## Features
- **Pure Go**: No external dependencies like LibreOffice or Docker required.
- **Embedded Font**: Includes a default Arial font for Excel rendering, works out of the box.
- **Excel support**: Custom robust renderer using `excelize` and `gopdf`.
- **Word support**: Uses `docx2pdf` for DOCX conversion.
- **Image support**: Convert single or multiple JPG and PNG images to a single PDF.
- **Simple API**: Easy to integrate into any project.

## Installation

```bash
go get github.com/kuyaninja/office-to-pdf
```

## Usage

### Converting Excel (.xlsx)

Excel conversion uses an embedded Arial font by default for proper rendering.

```go
package main

import (
    "github.com/kuyaninja/office-to-pdf"
    "log"
)

func main() {
    // Initialize converter (uses embedded font by default)
    converter := office_to_pdf.NewConverter("", "")

    err := converter.ConvertXlsxToPdf("input.xlsx", "output.pdf")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Converting Multiple Images to 1 PDF

```go
package main

import (
    "github.com/kuyaninja/office-to-pdf"
    "log"
)

func main() {
    converter := office_to_pdf.NewConverter("", "")

    images := []string{"page1.jpg", "page2.png", "page3.jpg"}
    err := converter.ConvertImagesToPdf(images, "combined.pdf")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Converting Word (.docx)

```go
package main

import (
    "github.com/kuyaninja/office-to-pdf"
    "log"
)

func main() {
    converter := office_to_pdf.NewConverter("", "")

    err := converter.ConvertDocxToPdf("input.docx", "output.pdf")
    if err != nil {
        log.Fatal(err)
    }
}
```

## CLI Tool

You can also use the included CLI tool:

```bash
go build -o convert cmd/convert/main.go

# Convert multiple images
./convert -output combined.pdf photo1.jpg photo2.png photo3.jpg

# Convert single files
./convert data.xlsx
./convert document.docx
```

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
