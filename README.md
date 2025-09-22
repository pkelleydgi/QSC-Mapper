# QSC to Q360 Data Mapper

A web-based tool for converting QSC pricing files to Q360 Load Sheet format.

## Features
- Drag-and-drop file upload
- Automatic column detection
- Instant processing
- Clean, professional interface

## Usage
1. Visit the [live app](https://YOUR_USERNAME.github.io/qsc-mapper/)
2. Upload your QSC pricing Excel file
3. Click "Process File"
4. Download the formatted Q360 file

## Mapping Rules
- SALES PART → MASTERNO & PARTNO
- LONG DESCRIPTION → DESCRIPTION
- NET DEALER → STANDARDCOST
- List Price → MSRP
- Fixed values: MANUFACTURER="QSC", TAXABLE="Y", USETAXFLAG="Y"

## Python Script
A command-line version is available in the `tools/` folder for batch processing.
