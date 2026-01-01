# Tamil Nadu Electoral Roll Voter Counter Application

## Overview

A Python GUI application that processes Tamil Nadu Electoral Roll PDFs to:
1. Count voters by gender (Male, Female, Third Gender)
2. Extract individual voter cards as JPEG images
3. Convert voter data to Excel with automatic error correction

---

## Features

### 1. Voter Count Statistics
- Extracts voter counts from the summary page of Electoral Roll PDFs
- Displays Male (ஆண்), Female (பெண்), Third Gender (மூன்றாம் பாலினம்), and Total (மொத்தம்) counts
- Uses OCR with Tamil + English language support

### 2. Voter Card Extraction
- Extracts individual voter cards from PDF pages 4 to n-1 (skips cover pages and summary)
- Each page contains a 3x10 grid (30 voter cards per page)
- Saves each card as a numbered JPEG file (1.jpg, 2.jpg, etc.)
- Automatically skips empty slots using brightness detection

### 3. Excel Generation with Auto-Fix
- Converts each voter card JPEG to a row in Excel
- **Columns**: S.No, Voter ID, Name, Relation Type, Relation Name, House No, Age, Gender
- **Two-pass processing**:
  - First pass: Standard OCR for all cards
  - Second pass: Enhanced OCR to fix missing data automatically

---

## Technical Implementation

### Dependencies
```python
pymupdf          # PDF rendering
pytesseract      # OCR engine (requires Tesseract installed)
pillow           # Image processing
openpyxl         # Excel file creation
tkinter          # GUI framework (built-in)
```

### Key Functions

#### `ocr_page(page, zoom=3)`
Renders a PDF page at high resolution and performs OCR.

#### `parse_voter_card(text)`
Extracts structured data from OCR text using regex patterns:
- **Voter ID**: Handles OCR variations like `IBU`, `1BU`, `18ப`, `AWP`
- **Name**: Extracts from `பெயர்:` field
- **Relation**: Detects Father (`தந்தை`) or Husband (`கணவர்`)
- **House No**: Extracts from `வீட்டு எண்:` field
- **Age/Gender**: Extracts from `வயது:` and `பாலினம்:` fields

#### `clean_ocr_text(text)`
Removes common OCR artifacts:
- "Photo is" / "available" text fragments
- Leading/trailing punctuation
- Multiple spaces

#### `enhanced_ocr_voter_card(image_path)`
Multi-approach OCR for difficult images:
1. Original image
2. High contrast (2x enhancement)
3. Grayscale + sharpness
4. Binary threshold (black/white)
5. 2x scaled image

Merges results from all approaches to maximize data extraction.

#### `fix_missing_data(ws, output_path)`
Automatic error correction:
1. Scans Excel for rows with missing fields
2. Finds corresponding JPEG by S.No
3. Re-OCRs with enhanced preprocessing
4. Updates Excel with extracted values

---

## PDF Structure (Tamil Nadu Electoral Roll)

```
Page 1-3:   Cover pages, signatures, maps (skipped)
Page 4-n:   Voter data pages
            - 3 columns x 10 rows = 30 cards per page
            - Each card contains voter details
Last Page:  Summary table with totals (used for counting)
```

### Voter Card Layout
```
┌─────────────────────────────────┐
│ [Serial No]     [Voter ID]      │
│ பெயர்: [Name]                   │
│ தந்தையின் பெயர்: [Father Name]  │
│ வீட்டு எண்: [House No]          │
│ வயது: [Age]  பாலினம்: [Gender]  │
│              [Photo]            │
└─────────────────────────────────┘
```

---

## Usage

### Running the Application
```bash
cd /mnt/d/search_my_name
python voter_counter_app.py
```

### Workflow
1. Click **Browse** to select an Electoral Roll PDF
2. Click **Count Voters** to get gender-wise statistics
3. Click **Extract Voter Cards** to:
   - Select output directory
   - Extract all voter cards as JPEGs
   - Generate Excel file with voter data
   - Automatically fix missing data

### Output Files
```
output_directory/
├── 1.jpg                    # First voter card
├── 2.jpg                    # Second voter card
├── ...
├── n.jpg                    # Last voter card
└── [pdf_name]_excel.xlsx    # Excel with all voter data
```

---

## OCR Challenges & Solutions

| Challenge | Solution |
|-----------|----------|
| Tamil text recognition | Use `lang='tam+eng'` in pytesseract |
| Voter ID misread as `18ப`, `1BU` | Multiple regex patterns for variations |
| "Photo is available" in House No | `clean_ocr_text()` removes artifacts |
| Missing data in some cards | Enhanced OCR with 5 preprocessing methods |
| Empty card slots | Brightness threshold detection (>252 = empty) |

---

## File Structure

```
/mnt/d/search_my_name/
├── voter_counter_app.py           # Main application
├── voter_counter_app_documentation.md  # This file
└── extract_voters/                # Output directory (created on extraction)
    ├── *.jpg                      # Voter card images
    └── *_excel.xlsx               # Generated Excel file
```

---

## Future Improvements

- [ ] Add batch processing for multiple PDFs
- [ ] Export to CSV format
- [ ] Add search/filter functionality in Excel
- [ ] Support other state electoral roll formats
- [ ] Add progress bar for fix pass
