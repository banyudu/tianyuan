# Excel Data Extraction Plan

This document outlines the comprehensive analysis and extraction rules for converting the parsed Excel JSON file to the four target data types: 子目信息, 含量表, 工作内容, 附注信息.

## Data Source Analysis

### File Structure
- **Input**: `output/parsed-excel.json` (2,869 cells with complete style information)
- **Source Excel**: `sample/input.xlsx` (839 rows × 32 columns)
- **Target outputs**: 4 CSV files matching `sample/output/` format

## Detection Rules & Patterns

### 1. Table Detection Rules

#### Primary Indicators
- **Medium borders** (`borderStyles.style: "medium"`) - primary table boundary indicator
- **Rectangular regions** with consistent border patterns
- **Table range pattern**: A73:AE80 (columns A to AE, variable row ranges)

#### Table Identification Process
1. Scan for cells with medium borders
2. Group adjacent cells with borders into rectangular regions
3. Identify table headers by spaced Chinese text patterns
4. Determine table boundaries by border consistency

### 2. Title Hierarchy Detection

#### Chapter Titles (第x章)
- **Content pattern**: `/^第[一二三四五六七八九十\d]+章/`
- **Location example**: Row 62 - "第一章 机械设备安装工程"
- **Style markers**:
  - Font: SimHei, size 10
  - Alignment: Center horizontal, middle vertical
  - Merge: Wide range (cols 1-28)
  - Borders: None

#### Section Titles (第x节)
- **Content pattern**: `/^第[一二三四五六七八九十\d]+节/`
- **Location example**: Row 63 - "第一节 减振装置安装"
- **Style markers**:
  - Font: SimHei, size 10
  - Alignment: Center horizontal, middle vertical
  - Merge: Wide range (cols 1-28)
  - Borders: None

#### Sub-section Titles
- **Content pattern**: `/^[一二三四五六七八九十]+、/`
- **Location example**: Row 71 - "一、减振装置安装"
- **Style markers**:
  - Font: SimHei, size 10
  - Alignment: Center horizontal, middle vertical
  - Merge: Widest range (cols 1-31)
  - Borders: None

### 3. Table Structure Analysis

#### Example Table A73:AE80

**Row 72**: Unit information ("台")
**Row 73**: Table headers and quota codes (1B-1, 1B-2, 1B-3, 1B-4, 1B-5, 1B-6)
**Row 74**: Sub-item base names ("钢架焊接减振台座(kg以内)")
**Row 75**: Amounts ("50")
**Rows 77-80**: Resource consumption data
- Row 77: 人工 (Labor)
- Row 78: 材料 (Materials)  
- Row 79: 机械 (Machinery)
- Row 80: Additional resources

#### Resource Data Structure
- **Resource names**: Column B (B77:B80)
- **Units**: Column J (J77:J80)
- **Consumption amounts**: Columns M-AE (aligned with quota codes)

## Data Extraction Rules

### 1. 子目信息 (Sub-item Information)

#### Hierarchical Structure
```
$ (Chapter) - 第x章
$$ (Section) - 第x节  
$$$ (Sub-section) - x、
$$$$ (Sub-sub-section) - Additional levels
子目项 (Items) - Quota codes with data
```

#### Sub-item Name Construction
- **Formula**: `${baseNameFromM74} ${amountFromM75}${unitFromA72}`
- **Example**: "钢架焊接减振台座(kg以内) 50台"

#### Numeric Fields Extraction
From table rows, extract values for:
- 基价 (Base price)
- 人工 (Labor cost)
- 材料 (Material cost)
- 机械 (Machinery cost)
- 管理费 (Management fee)
- 利润 (Profit)
- 其他 (Other costs)

### 2. 含量表 (Material/Resource Table)

#### Data Sources
- **Location**: Resource consumption rows (A77:AE80 in example table)
- **Resource types**: 人工=1, 材料=2, 机械=3, 其他=5

#### Field Mapping
- **编号**: Quota codes from table headers
- **名称**: Resource names from column B
- **规格**: Specifications (if available)
- **单位**: Units from column J
- **单价**: Unit prices (if available)
- **含量**: Consumption amounts from columns M-AE
- **主材标记**: "*" indicates main material
- **材料类别**: 1=人工, 2=材料, 3=机械, 5=其他
- **是否有明细**: "是" indicates detailed breakdown available

### 3. 工作内容 (Work Content)

#### **UPDATED RULE**: Location and Structure
- **Location**: **Same row as unit information (row above table)**
- **Example**: Row 72 contains both unit ("台") AND work content
- **Content pattern**: Contains "工作内容：" followed by description
- **Format**: Multiple quota codes can share same work content

#### Extraction Process
1. Find table unit row (row above table headers)
2. Check if cell contains work content information
3. Extract quota codes from table headers below
4. Map work content to relevant quota codes
5. **Consolidation**: Merge rows with identical work content values

#### Output Format
```csv
编号,工作内容
"1B-1,1B-2,1B-3","工作内容描述"
```

### 4. 附注信息 (Note Information)

#### Content Patterns
- **Pattern**: Contains "注 : 未包括" or "注: 未包括"
- **Location**: Scattered throughout document
- **Format**: Single quota code per note typically

#### Extraction Process
1. Scan all cells for note patterns
2. Extract associated quota codes (nearby cells or context)
3. **Consolidation**: Group notes with identical content
4. Format as comma-separated quota codes

#### Output Format
```csv
编号,附注信息
"1B-1","未包括电杆、地横木。"
```

## Implementation Strategy

### Phase 1: Table Detection
1. Scan parsed JSON for medium border patterns
2. Group cells into table regions
3. Identify table boundaries and headers
4. Extract quota code sequences

### Phase 2: Hierarchy Building
1. Detect chapter/section/sub-section titles
2. Build hierarchical structure with proper symbols
3. Associate tables with their parent sections

### Phase 3: Data Extraction
1. **子目信息**: Build hierarchical structure + extract sub-items
2. **含量表**: Extract resource consumption data from table rows
3. **工作内容**: Extract from unit rows, consolidate duplicates
4. **附注信息**: Scan for note patterns, consolidate duplicates

### Phase 4: Output Generation
1. Generate CSV files matching sample format exactly
2. Ensure proper escaping and formatting
3. Validate against reference output files

## Key Technical Considerations

### Style-Based Detection
- **Font patterns**: SimHei for headers, SimSun for content
- **Border weights**: Medium for boundaries, thin for internal
- **Alignment**: Center for titles, varied for content
- **Merge patterns**: Wider merges indicate higher hierarchy

### Data Consolidation Rules
- **工作内容**: Group by identical content, join quota codes with commas
- **附注信息**: Group by identical content, join quota codes with commas
- **含量表**: One row per quota code + resource combination
- **子目信息**: Hierarchical structure with proper symbol prefixes

### Validation Criteria
- Compare output row counts with reference files
- Verify quota code patterns and sequences
- Check hierarchical symbol consistency
- Validate CSV formatting and escaping

## Expected Output Structure

### 子目信息.csv
```csv
,定额号,子目名称,基价,人工,材料,机械,管理费,利润,其他,图片名称
$,,"第一章 机械设备安装工程",0,0,0,0,0,0,0,
$$,,"第一节 减振装置安装",0,0,0,0,0,0,0,
$$$,,"一、减振装置安装",0,0,0,0,0,0,0,
,1B-1,"钢架焊接减振台座(kg以内) 50台",基价值,人工值,材料值,机械值,管理费值,利润值,其他值,
```

### 含量表.csv
```csv
编号,名称,规格,单位,单价,含量,主材标记,材料号,材料类别,是否有明细
1B-1,综合用工二类,,工日,0,含量值,,1,
```

### 工作内容.csv  
```csv
编号,工作内容
"1B-1,1B-2,1B-3","具体工作内容描述"
```

### 附注信息.csv
```csv
编号,附注信息  
"1B-1","未包括相关项目说明"
```