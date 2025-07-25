# Excel File Conversion Tool Plans

## Project Overview
A TypeScript tool to convert Excel files from the format "建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx" to a structured directory format with 3 separate Excel files.

## Input Analysis
- **Input File**: `data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx`
- **Structure**: Single worksheet with 839 rows, 32 columns (A-AF)
- **Content**: Complex hierarchical data with merged cells representing:
  - Item codes (e.g., "1B-1", "1B-2")
  - Chinese descriptions and specifications
  - Units of measurement (个, m, kg, etc.)
  - Quantities and measurements
  - Work content and notes

## Output Requirements
Convert to directory: `data/补充子目/` containing:

1. **子目信息.xls** - Subitem information table
2. **工作内容、附注信息表.xls** - Work content and notes table  
3. **含量表.xls** - Content/material consumption table

## Technical Approach

### Phase 1: Enhanced Analysis
- [ ] Fix cell value parsing to handle rich text and formulas
- [ ] Correctly identify merged cell regions and their hierarchical structure
- [ ] Map data patterns to understand table organization
- [ ] Identify header rows and data boundaries
- [ ] Create data model for the three output table types

### Phase 2: Data Extraction Logic
- [ ] Implement merged cell parsing for hierarchical data
- [ ] Extract subitem information (codes, names, specifications)
- [ ] Extract work content and notes
- [ ] Extract material consumption data with quantities and units
- [ ] Handle variations in merged cell patterns across different sections

### Phase 3: Output Generation
- [ ] Create structured data models for each output file
- [ ] Generate Excel files in the required format (.xls)
- [ ] Ensure proper table headers and formatting
- [ ] Validate output against reference files

### Phase 4: Robustness and Future-Proofing
- [ ] Handle varying merged cell patterns
- [ ] Add configuration for different input formats
- [ ] Error handling and validation
- [ ] Logging and debugging capabilities

## Implementation Details

### Key Libraries
- `exceljs` - Excel file parsing and generation
- `xlsx` - Additional Excel format support (if needed)

### Data Structures
```typescript
interface SubitemInfo {
  code: string;
  name: string;
  unit: string;
  specification?: string;
}

interface WorkContent {
  code: string;
  workContent: string;
  notes?: string;
}

interface MaterialContent {
  subitemCode: string;
  materialCode: string;
  materialName: string;
  unit: string;
  quantity: number;
}
```

### Processing Pipeline
1. **Parse Input** → Read Excel file and analyze structure
2. **Extract Data** → Parse merged cells and extract structured data
3. **Transform** → Organize data into three categories
4. **Generate Output** → Create three separate Excel files

## Risk Mitigation
- **Merged Cell Variations**: Create flexible parsing logic that can adapt to different merge patterns
- **Data Format Changes**: Implement pattern-based recognition rather than hardcoded column positions
- **Output Format Compatibility**: Ensure output files match expected format exactly

## Success Criteria
- [ ] Successfully parse the input Excel file structure
- [ ] Extract all three types of data accurately
- [ ] Generate output files matching the reference format
- [ ] Handle edge cases and data variations
- [ ] Create maintainable and extensible code 
