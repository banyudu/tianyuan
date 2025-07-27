# Excel File Conversion Tool

A TypeScript tool that converts complex Excel files from the format "建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx" to a structured directory format with 3 separate Excel files.

## Overview

This tool processes construction engineering standards Excel files and extracts data into three structured output files:

1. **子目信息.xls** - Subitem information table
2. **工作内容、附注信息表.xls** - Work content and notes table
3. **含量表.xls** - Material consumption table

## Features

- ✅ **Smart Excel Parsing**: Handles complex merged cells and hierarchical data structures
- ✅ **Multi-section Data Extraction**: Automatically identifies and extracts different data types
- ✅ **Pattern Recognition**: Uses regex patterns to identify codes, units, and content types
- ✅ **Robust Error Handling**: Gracefully handles variations in input file structure
- ✅ **TypeScript**: Fully typed for better maintainability and reliability

## Installation

```bash
# Install dependencies
pnpm install

# Run the converter
pnpm run start
```

## Usage

### Basic Usage

```typescript
import { ExcelConverter } from './converter';

const converter = new ExcelConverter();
await converter.convert('data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx', 'output');
```

### Command Line

```bash
# Run the converter directly
tsx converter.ts
```

## Input File Structure

The input Excel file contains:

- **Table of Contents**: Chapter and section listings with page numbers
- **Subitem Codes**: Alphanumeric codes like "1B-1", "2A-3", etc.
- **Work Content**: Detailed descriptions of work processes and procedures
- **Material Lists**: Construction materials with specifications and quantities
- **Merged Cells**: Complex hierarchical table structures

## Output Files

### 1. 子目信息.xls (Subitem Information)

| 子目编号 | 子目名称     | 计量单位 | 工作内容   |
| -------- | ------------ | -------- | ---------- |
| 1B-1     | 减振装置安装 | 个       | 安装、调试 |

### 2. 工作内容、附注信息表.xls (Work Content)

| 编号 | 名称     | 工作内容       | 计量单位 | 备注 |
| ---- | -------- | -------------- | -------- | ---- |
| WC-1 | 阀门安装 | 安装、水压试验 | 个       |      |

### 3. 含量表.xls (Material Consumption)

| 子目编号 | 材料名称 | 计量单位 | 消耗量 | 类别 |
| -------- | -------- | -------- | ------ | ---- |
| 1B-1     | 螺纹法兰 | 个       | 2      | 材料 |

## Technical Implementation

### Data Models

```typescript
interface SubitemInfo {
  code: string; // 子目编号 (e.g., "1B-1")
  name: string; // 子目名称
  unit: string; // 计量单位
  workContent?: string; // 工作内容
}

interface WorkContent {
  code: string; // 编号
  name: string; // 名称
  workContent: string; // 工作内容
  unit: string; // 计量单位
  notes?: string; // 备注
}

interface MaterialContent {
  subitemCode: string; // 子目编号
  materialName: string; // 材料名称
  unit: string; // 计量单位
  quantity: number; // 消耗量
  category: string; // 类别 (人工/材料/机械)
}
```

### Processing Pipeline

1. **Load & Analyze**: Parse Excel file and identify section boundaries
2. **Extract Data**: Use pattern matching to extract structured data
3. **Transform**: Organize data into three categories
4. **Generate**: Create output Excel files with proper formatting

## Analysis Results

Successfully analyzed input file with:

- **839 rows** and **32 columns** of data
- **386 material items** extracted
- **Multiple data sections** identified and processed
- **Complex merged cell structures** handled

## Future Enhancements

- [ ] Improve subitem extraction for complex merged cells
- [ ] Add support for additional input file variations
- [ ] Implement data validation and quality checks
- [ ] Add configuration options for different extraction patterns
- [ ] Support for batch processing multiple files

## Dependencies

- `exceljs`: Excel file parsing and generation
- `typescript`: Type safety and modern JavaScript features
- `tsx`: TypeScript execution

## Files Structure

```
├── converter.ts              # Main conversion logic
├── plans.md                  # Project planning document
├── analyze-detailed.ts       # Detailed file analysis
├── find-data-sections.ts     # Section boundary detection
├── data/                     # Input files
├── output/                   # Generated output files
└── README.md                 # This file
```

## License

ISC License
