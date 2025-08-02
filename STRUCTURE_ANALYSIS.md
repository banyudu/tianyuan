# Excel File Structure Analysis

This document provides a detailed breakdown of the `sample/input.xlsx` file structure based on analysis. This is a construction cost estimation document converted from PDF with OCR processing.

## File Overview
- **Total Rows**: 839
- **Total Columns**: 32 (A to AF)
- **Original Format**: PDF converted to Excel via OCR
- **Content Type**: Construction cost estimation with detailed material and labor breakdowns

## File Structure Breakdown

### 1. Document Header and Table of Contents (Rows 1-61)
- **Rows 1-3**: Document title and headers
  - Row 1: "建设工程消耗量标准及计算规则（安装工程）" (Construction Engineering Consumption Standards and Calculation Rules - Installation Engineering)
  - Row 2: "补充子目" (Supplementary Sub-items)
  - Row 3: "目录" (Table of Contents)
- **Rows 4-61**: Table of contents listing chapters and sections

### 2. Chapter/Section Structure (Starting Row 62)
- **Row 62**: "第一章 机械设备安装工程" (Chapter 1: Mechanical Equipment Installation Engineering)
- **Row 63**: "第一节 减振装置安装" (Section 1: Vibration Reduction Device Installation)
- **Row 65**: "一、本节项目包括下列内容：" (1. This section includes the following content:)

### 3. Data Tables Structure (Starting Row 73)

#### Table 1: A73:AF80 - First Sub-item Table
**Header Structure:**
- **Row 73**: "子目编号" (Sub-item Code) headers spanning columns A-L, then quota codes:
  - Columns M-P: "1B-1" 
  - Columns Q-T: "1B-2"
  - Columns U-W: "1B-3" 
  - Columns X-Y: "1B-4"
- **Row 74**: "子目名称" (Sub-item Name) - All show "钢架焊接减振台座(kg以内)" (Steel Frame Welded Vibration Platform (within kg))
- **Row 75**: Capacity specifications - 50, 100, 200, 300 kg respectively
- **Row 76**: "人材机名称" (Labor/Material/Machine Name) and "消耗" (Consumption) headers

**Data Rows:**
- **Row 77**: 人工 (Labor) - "综合用工二类" (Comprehensive Labor Class 2) with consumption values
- **Row 78-79**: 材料 (Materials) - Including steel frame, oxygen, acetylene gas, welding rods
- **Row 80**: 机械 (Machinery) - "交流电焊机 21kV·A" (AC Welding Machine 21kV·A)

#### Table 2: Starting Row 82 - Second Sub-item Table  
- **Row 82**: Next set of quota codes:
  - Columns P-U: "1B-7"
  - Columns V-Y: "1B-8"
- **Row 83**: "钢筋混凝土减振台座(m3以内)" (Reinforced Concrete Vibration Platform (within m3))
- **Row 84**: Capacity specifications - 0.2 and 0.5 m3

## Quota Code Pattern Analysis

### Pattern Recognition
- **Format**: `\d+[A-Z]-\d+` (e.g., 1B-1, 1B-2, 1B-3, etc.)
- **Location**: Quota codes appear in table headers starting from column M onwards
- **Series Found**: 1B-1, 1B-2, 1B-3, 1B-4, 1B-7, 1B-8

### Table Structure Pattern
Each table follows this structure:
1. **Row 1**: Quota codes spanning multiple columns (merged cells)
2. **Row 2**: Sub-item names (材料名称)
3. **Row 3**: Specifications/capacity values  
4. **Row 4**: Headers for "人材机名称" and "消耗"
5. **Subsequent rows**: Detailed breakdown of:
   - 人工 (Labor)
   - 材料 (Materials) 
   - 机械 (Machinery)

## Key Pattern Locations Discovered

### Work Content Patterns
- **Row 763**: "一、工作内容：" (Work content section start)  
- **Row 778**: "一、工作内容："
- **Row 819**: "一、工作内容："

### Note Information Patterns  
- **Row 301**: "注 : 未包括电杆、地横木。"
- **Row 310**: "注 : 未包括木电杆、水泥接腿杆、地横木、圆木、连接铁件及螺栓。"
- **Row 323**: "注 : 未包括撑杆、圆木、连接铁件及螺栓。"
- **Row 342**: "注 : 未包括横担、绝缘子、连接铁件及螺栓。"

## Expected Output Format Analysis

Based on sample output files, the converter must produce exactly these 4 CSV files:

### 1. 子目信息.csv (Sub-item Information)
**Headers**: `,定额号,子目名称,基价,人工,材料,机械,管理费,利润,其他,图片名称,`

**Structure Pattern**:
- `$` = Chapter level (e.g., "$,第一章, 机械设备安装工程")
- `$$` = Section level (e.g., "$$,第一节,减振装置安装")  
- `$$$` = Sub-section level (e.g., "$$$,一、,减振装置安装")
- No prefix = Quota codes (e.g., ",1B-1,钢架焊接减振台座(kg以内) 50&台,0,0,0,0,0,0,0,,")

**Key Observations**:
- Hierarchical structure with $, $$, $$$ symbols
- Quota codes follow pattern: 1B-1, 1B-2, etc.
- Names include capacity with `&` separator (e.g., "50&台", "100&台")
- All numeric fields (基价,人工,材料,机械,管理费,利润,其他) are set to 0
- Empty columns for 图片名称

### 2. 含量表.csv (Material Content Table)
**Headers**: `编号,名称,规格,单位,单价,含量,主材标记,材料号,材料类别,是否有明细,,,`

**Data Pattern**:
- Each quota code (编号) has multiple material entries
- **材料类别** values: `1`=人工, `2`=材料, `3`=机械, `5`=other  
- **主材标记**: `*` for main materials, empty otherwise
- **含量** = consumption quantity
- Multiple rows per quota code for different materials

**Example**:
```
1B-1,综合用工二类,,工日,0.122,0.122,,,1,,,,
1B-1,钢架焊接减振台座,,台,1,1,*,,5,,,,
1B-1,氧气,,m3,0.016,0.016,,,2,,,,
```

### 3. 工作内容.csv (Work Content)
**Headers**: `编号,工作内容`

**Data Pattern**:
- **编号**: Comma-separated quota codes in quotes (e.g., "2B-84,2B-85,2B-86")
- **工作内容**: Detailed work description
- Each row represents one work content entry for multiple related quota codes

**Example**:
```
"2B-84,2B-85,2B-86",测位、划线、支架安装、吊装灯杆、组装接线、接地。
```

### 4. 附注信息.csv (Note Information)  
**Headers**: `编号,附注信息`

**Data Pattern**:
- **编号**: Comma-separated quota codes in quotes 
- **附注信息**: Note text (usually exclusions starting with "未包括")
- Similar format to work content but for notes/exclusions

**Example**:
```
"2B-9,2B-10,2B-11,2B-12,2B-13,2B-14,2B-15",未包括电杆、地横木。
```

## Implementation Strategy for Data Extraction

### 1. 子目信息 (Sub-item Information) Extraction
**Source**: Table structures starting around row 73
- Scan for quota code patterns in table headers (rows like 73, 82)
- Extract quota codes from merged cells (e.g., M73:P73 = "1B-1")
- Get sub-item names from subsequent rows (e.g., row 74)  
- Get capacity values (e.g., row 75: 50, 100, 200, 300)
- Build hierarchical structure with chapter/section info from rows 62-63

### 2. 含量表 (Material Content Table) Extraction  
**Source**: Material data in table rows 77-80+
- For each quota code, extract material consumption data
- Row 77: 人工 data (Material category = 1)
- Rows 78-79: 材料 data (Material category = 2)  
- Row 80: 机械 data (Material category = 3)
- Map material names, units, and consumption quantities
- Mark main materials with "*" in 主材标记

### 3. 工作内容 (Work Content) Extraction
**Source**: Rows 763, 778, 819 and nearby rows
- Search for "一、工作内容：" pattern
- Extract quota code groups and their work descriptions
- Look in nearby rows for the actual work content text
- Group related quota codes together with shared work descriptions

### 4. 附注信息 (Note Information) Extraction
**Source**: Rows 301, 310, 323, 342 and similar patterns
- Search for "注 : 未包括" pattern throughout document
- Extract associated quota codes (likely in nearby rows/cells)
- Group quota codes that share the same note information
- Format as comma-separated quoted strings

## Critical Implementation Requirements

### Data Format Matching
- **子目信息**: Must use exact hierarchical symbols ($, $$, $$$)
- **含量表**: Must correctly assign material categories (1,2,3,5)
- **工作内容**: Quota codes must be comma-separated in quotes
- **附注信息**: Same format as work content but for notes

### Table Detection Strategy
- Tables start with quota code patterns in headers
- Each table can span multiple quota codes horizontally
- Tables are separated by unit specification rows (e.g., "单位：块")
- Merged cells require careful handling for quota code extraction

### Pattern Matching Locations
- **Quota codes**: Primarily in table headers (rows 73+, 82+)
- **Work content**: Specific rows 763, 778, 819
- **Notes**: Scattered throughout (301, 310, 323, 342)
- **Chapters/Sections**: Rows 62-63 for hierarchical structure

## File Locations for Reference
- **Input**: `sample/input.xlsx` (839 rows × 32 columns)
- **Analysis Script**: `analyze_structure.ts` 
- **Expected Output**: `sample/output/` (4 CSV files)
- **Current Output**: `output/csv/` and `output/excel/`
- **Structure Documentation**: `STRUCTURE_ANALYSIS.md` (this file)