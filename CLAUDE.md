# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an Excel converter tool that transforms OCR-generated Excel files from WPS into standardized CSV and Excel output files. The project specifically handles PDF documents that have been OCR-processed and contain complex formatting with merged cells and border information.

**Input**: OCR-processed Excel files with merged cells and border formatting (sample/input.xlsx)
**Output**: Four standardized CSV/Excel files (子目信息, 含量表, 工作内容, 附注信息)

## Core Architecture

### Main Components

1. **ExcelAnalyzer** (`src/excel-analyzer.ts`): Original analyzer with border detection and region identification
2. **ImprovedExcelConverter** (`convert.ts`): New improved converter with pattern-based extraction
3. **DataExtractor** (`src/data-extractor.ts`): Legacy data extraction logic
4. **FileExporter** (`src/file-exporter.ts`): Exports data to CSV and Excel formats
5. **Types** (`src/types.ts`): TypeScript interfaces defining all data structures

### Data Processing Flow

The improved converter (`convert.ts`) uses a pattern-based approach:

1. Load Excel file using ExcelJS library
2. Scan for specific patterns:
   - Quota codes: `/^\d+[A-Z]-\d+$/` (e.g., "1B-1", "2B-84")
   - Chapter headers: `/^第[一二三四五六七八九十]+章/`
   - Work content: Contains "工作内容" and "："
   - Notes: Contains "注 : 未包括" or "注: 未包括"
3. Extract four data types using targeted pattern matching
4. Export to CSV format matching sample output

## Development Commands

```bash
# Install dependencies
npm install

# Run original converter
npm start

# Run improved converter  
npx tsx convert.ts

# Format code
npm run format

# Debug file structure
npx tsx src/debug-analyzer.ts
```

## File Structure and Key Locations

### Input/Output
- Input: `sample/input.xlsx` (839 rows, 32 columns)
- Expected output: `sample/output/` (reference CSV files)
- Current output: `output/csv/` and `output/excel/`
- Improved output: `output_improved/`

### Key Row Locations in Input File
- Chapter headers start around row 62 ("第一章 机械设备安装工程")
- Quota codes begin around row 73
- Work content patterns around rows 763, 778, 819
- Note information scattered throughout (rows 301, 310, 323, etc.)

## Data Type Recognition Patterns

### 工作内容 (Work Content)
- Pattern: Contains "工作内容：" followed by detailed work descriptions
- Format: "编号,工作内容" (comma-separated quota codes, work description)
- Sample: "2B-84,2B-85,2B-86","测位、划线、支架安装、吊装灯杆、组装接线、接地。"

### 附注信息 (Note Information)  
- Pattern: Contains "注 : 未包括" followed by exclusion details
- Format: "编号,附注信息" (single quota code, note text)
- Sample: "2B-16","未包括电杆、地横木。"

### 子目信息 (Sub-item Information)
- Hierarchical structure with symbols: $, $$, $$$, $$$$
- Includes chapter/section headers and quota code entries
- Format: Symbol, 定额号, 子目名称, and numeric fields (基价, 人工, 材料, etc.)

### 含量表 (Material Table)
- Material consumption data with categories
- Format: 编号, 名称, 规格, 单位, 单价, 含量, 主材标记, 材料号, 材料类别, 是否有明细
- Material categories: 1=人工, 2=材料, 3=机械, 5=other

## Important Implementation Notes

### Pattern Matching Strategy
- Use `getCellValue()` helper to handle Excel object types (richText, formulas, etc.)
- Search in row ranges rather than cell-by-cell for efficiency
- Look in nearby rows for associated data (quota codes often appear near descriptions)

### Data Cleaning
- Trim whitespace and handle empty cells
- Remove duplicate text patterns
- Handle merged cells by getting master cell values
- Convert boolean values: 主材标记 uses "*", 是否有明细 uses "是"

### CSV Export Format
- Headers match exactly with sample files
- Empty columns ("", 图片名称) preserved for compatibility
- Proper CSV escaping for values containing commas or quotes

## Debugging and Analysis Tools

- `src/debug-analyzer.ts`: Analyzes input file structure and patterns
- Shows row-by-row breakdown of first 50 rows
- Identifies quota codes, chapters, work content, and note locations
- Essential for understanding new input file formats

## Testing and Validation

Compare outputs with reference files in `sample/output/`:
- Check quota code extraction accuracy
- Verify hierarchical structure in 子目信息
- Ensure CSV formatting matches exactly
- Validate material category assignments

## Dependencies

- **ExcelJS**: Primary Excel file processing library
- **tsx**: TypeScript execution for Node.js
- **fs/path**: Node.js file system operations
- **node-xlsx**, **xlsx**: Alternative Excel libraries (legacy support)