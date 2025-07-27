# Excel Conversion Fix Plan

## Problem Analysis

Based on the comparison with reference files, our current extraction has several critical issues:

### 🔍 **Reference File Structure Analysis**

#### 1. 子目信息.xls (337 rows)

- **Headers**: `定额号 | 子目名称 | 基价 | 人工 | 材料 | 机械 | 管理费 | 利润 | 其他 | 图片名称`
- **Structure**: Hierarchical with $, $$, $$$ markers for chapters
- **Data**: 336 rows with actual subitem codes (1B-1, 1B-2, etc.)
- **Format**: Each row is one subitem with complete pricing information

#### 2. 工作内容、附注信息表.xls (37 rows)

- **Headers**: `编号 | 工作内容`
- **Format**: Multiple subitem codes per row (e.g., "2B-84,2B-85,2B-86")
- **Content**: Work descriptions mapped to multiple related subitems

#### 3. 含量表.xls (2694 rows)

- **Headers**: `编号 | 名称 | 规格 | 单位 | 单价 | 含量 | 主材标记 | 材料号 | 材料类别 | 是否有明细`
- **Format**: One material per row, linked to subitem code
- **Data**: 2693 material consumption records

### 🚨 **Current Problems**

1. **Subitem Extraction**: Only found 6 codes, should find hundreds
2. **Wrong Headers**: Our headers don't match reference format
3. **Section Detection Failed**: Looking in wrong parts of input file
4. **Data Format Mismatch**: Structure completely different from reference
5. **Material Extraction**: Concatenating multiple items incorrectly

## 🛠️ **Fix Strategy**

### Phase 1: Correct Data Location Discovery

- [ ] Find the actual data tables in the input file (not table of contents)
- [ ] Identify where subitem codes are actually defined with pricing
- [ ] Locate material consumption tables with proper structure
- [ ] Map work content descriptions to correct subitem groups

### Phase 2: Header Format Correction

- [ ] Update 子目信息.xls headers to match reference: `定额号 | 子目名称 | 基价 | 人工 | 材料 | 机械 | 管理费 | 利润 | 其他 | 图片名称`
- [ ] Update 工作内容、附注信息表.xls headers: `编号 | 工作内容`
- [ ] Update 含量表.xls headers: `编号 | 名称 | 规格 | 单位 | 单价 | 含量 | 主材标记 | 材料号 | 材料类别 | 是否有明细`

### Phase 3: Data Extraction Logic Rewrite

- [ ] **Subitem Extraction**: Find hierarchical structure with pricing data
- [ ] **Work Content Extraction**: Group multiple subitem codes with shared work descriptions
- [ ] **Material Extraction**: Extract individual material records properly
- [ ] **Data Validation**: Ensure extracted data matches reference patterns

### Phase 4: Output Format Matching

- [ ] Generate hierarchical structure in 子目信息.xls (with $, $$, $$$ markers)
- [ ] Create grouped work content entries
- [ ] Generate individual material consumption records
- [ ] Validate row counts match expected scale (337, 37, 2694 respectively)

## 🎯 **Implementation Plan**

### Step 1: Deep Input File Analysis

Create new analysis to find actual data tables (not TOC sections)

### Step 2: Rewrite Extraction Logic

- Focus on finding tabular data with pricing information
- Look for material lists with quantities and specifications
- Identify work content groupings

### Step 3: Match Reference Structure

- Implement hierarchical output for subitems
- Group work content by related codes
- Generate one-material-per-row format

### Step 4: Validation & Testing

- Compare row counts with reference files
- Validate data structure and content quality
- Test with edge cases and variations

## 🚀 **Next Actions**

1. **Immediate**: Create new input file scanner to find actual data tables
2. **Priority**: Rewrite extraction logic based on reference file structure
3. **Validation**: Implement comparison tools to verify output quality
4. **Testing**: Ensure output matches reference file patterns exactly
