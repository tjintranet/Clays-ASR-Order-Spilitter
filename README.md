# ASR Daily Order Splitter

A browser-based tool for splitting ASR Daily Order Excel files into separate workbooks grouped by Cover Spec and Paper type. Cover specification codes are automatically decoded into plain-English descriptions.

## Features

- **Excel upload** — drag and drop or browse to load an `.xlsx` or `.xls` ASR Daily Order file
- **Three split modes** — group rows by Cover Spec + Paper, Cover Spec only, or Paper only
- **Spec code decoding** — Cover Spec codes (e.g. `C400P2`) are automatically translated into human-readable descriptions in the preview table and the Combined Workbook Summary sheet
- **Per-group download** — download any individual group as a standalone `.xlsx` file
- **Download All** — download every group as a separate `.xlsx` file in sequence
- **Combined Workbook** — download a single `.xlsx` containing a Summary sheet plus one tab per group
- **Clear** — reset the app and load a new file

## File Structure

```
├── index.html      # Markup and layout
├── style.css       # Styles (Bootstrap overrides + table rules)
├── script.js       # All application logic
└── README.md       # This file
```

## Usage

1. Open `index.html` in a modern web browser (no server required)
2. Click the file input or drag an `.xlsx` file onto it
3. Select a **Split Mode** using the radio buttons
4. Review the preview table — each row represents one output sheet
5. Click the **download** button on any row to save that group, or use the header buttons to download all at once

## Expected Excel Format

The uploaded file must contain at least the following column headers in the first sheet:

| Column | Description |
|--------|-------------|
| `Cover Spec` | Specification code (e.g. `C400P2`) |
| `Paper` | Paper stock code (e.g. `DHOL01`) |
| `GSM` | Paper weight |
| `Micron` | Paper thickness |

All other columns present in the file are preserved in the exported sheets.

## Cover Spec Code Structure

Codes are decoded character by character using the following structure:

| Position | Description |
|----------|-------------|
| 1st | Product type |
| 2nd | Outside colours |
| 3rd | Inside colours |
| 4th | Type of finish |
| 5th | Surface texture |
| 6th | Material weight |
| 7th+ | Special processes (optional, 1–2 character codes) |

### Product Types

| Code | Description |
|------|-------------|
| `C` | Cover |
| `W` | Cover with Flaps |
| `J` | Jacket |
| `T` | Tip-In |
| `F` | Cover For Case |

### Colour Configurations (positions 2 & 3)

| Code | Description |
|------|-------------|
| `0` | No Colour Print |
| `1` | 1 Spot Colour |
| `2` | 2 Spot Colours |
| `3` | 3 Spot Colours |
| `4` | 4 Process Colours |
| `5` | 4 Process Colours + 1 Spot Colour |
| `6` | 4 Process Colours + 2 Spot Colours |
| `7` | 4 Spot Colours |
| `8` | 4 Process Colours + 3 Spot Colours |
| `9` | 4 Process Colours + 4 Spot Colours |

### Finish Types (position 4)

| Code | Description |
|------|-------------|
| `0` | No Finish |
| `1` | Gloss Varnish In Line |
| `2` | Gloss Varnish In Line + Matt Varnish Offline |
| `3` | Gloss Varnish Off Line |
| `4` | Matt Varnish Off Line |
| `5` | Gloss Laminate (Standard) |
| `6` | Matt Laminate (Standard) |
| `7` | Matt Laminate (Standard) / Gloss Spot Varnish |
| `8` | Silk Laminate |
| `9` | Anti-Scuff Laminate |
| `A` | Gloss Laminate (Standard) / Matt Spot UV |
| `B` | Silk Laminate / Matt Spot UV |
| `C` | Anti-Scuff Laminate / Gloss Spot UV |
| `D` | Gloss Varnish Off Line + Matt Spot UV |
| `E` | Matt Varnish In Line + Gloss Spot UV |
| `F` | Matt Varnish In Line |
| `G` | Matt Varnish Off Line + Gloss Spot UV |
| `H` | Outwork Lamination |
| `J` | Outwork Lamination / Gloss Spot UV |
| `K` | Outwork Lamination / Matt Spot UV |
| `L` | Gloss Spot UV |
| `M` | Matt Spot UV |
| `N` | Gloss Varnish In Line + Matt Spot UV |
| `Q` | Soft Matt Lam |
| `R` | Soft Matt Lam / Gloss Spot Varnish |
| `V` | Recycled Matt Laminate |
| `W` | Recycled Matt Laminate / Gloss Spot Varnish |
| `Y` | Recycled Gloss Laminate |
| `Z` | Recycled Gloss Laminate / Matt Spot UV |

### Surface Texture (position 5)

| Code | Description |
|------|-------------|
| `P` | Plain |
| `G` | Grained |

### Material Weight (position 6)

| Code | GSM |
|------|-----|
| `1` | 220 gsm |
| `2` | 220 gsm |
| `3` | 260 gsm |
| `4` | 150 gsm |
| `5` | 135 gsm |
| `6` | 130 gsm |
| `7` | 220 gsm |

### Special Processes (positions 7+)

| Code | Description |
|------|-------------|
| `F` | Fluorescent |
| `S` | Spot Colour |
| `M` | Non-Conventional Metallic |
| `K` | Conventional Metallic (used with M) |
| `B` | Blocked (after print, before laminate) |
| `E` | Embossed |
| `D` | Debossed |
| `C` | Die-Cutting |
| `P` | Print Over Foil |
| `L` | Block Over Laminate |
| `U` | Uncoated Printing |
| `PB` | Print Black Over Foil |
| `BE` | Block & Emboss (same pass) |
| `DE` | Deboss & Emboss (same pass) |
| `BD` | Block & Deboss (same pass) |
| `S1` | Other Spot UV |
| `S2` | Pile Spot UV |
| `S3` | Glitter Spot UV |
| `V1` | Glow Varnish |
| `H1` | Holographic Lam |

## Exported File Format

All exported `.xlsx` files share the same formatting:

- Bold white header row with blue background
- Auto-fitted column widths
- Frozen header row for easy scrolling
- All original columns from the source file are preserved

The **Combined Workbook** includes an additional **Summary** sheet as the first tab, listing each group with its order count, Cover Spec code, decoded description, Paper, GSM, and Micron.

## Technical Details

| Item | Detail |
|------|--------|
| Technology | HTML5, CSS3, JavaScript (ES6+) |
| Styling | Bootstrap 5.3.2, Font Awesome 6.4.0 |
| Excel handling | SheetJS (xlsx) 0.18.5 |
| Dependencies | All libraries loaded via CDN |
| Server required | No — runs entirely in the browser |

## Browser Compatibility

Chrome 60+, Firefox 55+, Safari 12+, Edge 79+
