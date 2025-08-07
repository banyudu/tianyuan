import * as fs from 'fs';
import * as path from 'path';
import { CellData, ParsedExcelData } from './src/types';
import { 
  StructuredDocument, 
  Chapter, 
  Section, 
  SubSection, 
  TableArea, 
  TableRange, 
  TableStructure,
  NormInfo,
  ResourceInfo,
  ResourceConsumption,
  CellInfo, 
  BorderInfo 
} from './src/structured-types';

class StructuredExcelParser {
  private data: ParsedExcelData;
  private cellMap: Map<string, CellData>;
  private processedRows: Set<number>;

  constructor(jsonData: ParsedExcelData) {
    this.data = jsonData;
    this.cellMap = new Map();
    this.processedRows = new Set();
    
    // Create cell lookup map
    for (const cell of this.data.cells) {
      this.cellMap.set(`${cell.row}-${cell.col}`, cell);
      this.cellMap.set(cell.address, cell);
    }
  }

  private getCell(row: number, col: number): CellData | undefined {
    return this.cellMap.get(`${row}-${col}`);
  }

  private getCellValue(row: number, col: number): string {
    const cell = this.getCell(row, col);
    if (!cell || !cell.value) return '';
    return String(cell.value).trim();
  }

  private getBorderInfo(cell: CellData): BorderInfo {
    const borders = cell.borderStyles || {};
    return {
      hasTop: !!(borders.top?.style),
      hasBottom: !!(borders.bottom?.style),
      hasLeft: !!(borders.left?.style),
      hasRight: !!(borders.right?.style),
      topStyle: borders.top?.style,
      bottomStyle: borders.bottom?.style,
      leftStyle: borders.left?.style,
      rightStyle: borders.right?.style
    };
  }

  private isChapterTitle(value: string, cell?: CellData): boolean {
    const hasChapterPattern = /^第[一二三四五六七八九十\d]+章/.test(value);
    if (!cell || !hasChapterPattern) return hasChapterPattern;
    
    // Use font information to confirm chapter titles (SimHei font)
    return cell.font?.name === 'SimHei';
  }

  private isSectionTitle(value: string, cell?: CellData): boolean {
    // Only SimHei font cells can be section titles
    if (cell?.font?.name !== 'SimHei') {
      return false;
    }
    
    // True section titles like "第一节 减振装置安装"
    const hasSectionPattern = /^第[一二三四五六七八九十\d]+节/.test(value);
    return hasSectionPattern;
  }

  private isSubSectionTitle(value: string, cell?: CellData): boolean {
    // Only consider subsection titles that use SimHei font (structural headers)
    if (cell?.font?.name !== 'SimHei') {
      return false;
    }
    
    // Subsection titles with spaces like "一 、减振装置安装" use SimHei font  
    const hasSpacedSubSectionPattern = /^[一二三四五六七八九十]+\s+、/.test(value);
    if (hasSpacedSubSectionPattern) {
      return true;
    }
    
    // Numbered subsections like "(1)" - but these are rare and should also have SimHei
    return /^\(\d+\)/.test(value);
  }

  private isNormCode(value: string): boolean {
    return /^\d+[A-Z]-\d+$/.test(value);
  }

  private areInSameTable(quota1: {row: number; col: number}, quota2: {row: number; col: number}): boolean {
    // Check if two quota codes are in the same table by looking at the area between them
    const minRow = Math.min(quota1.row, quota2.row);
    const maxRow = Math.max(quota1.row, quota2.row);
    const minCol = Math.min(quota1.col, quota2.col);
    const maxCol = Math.max(quota1.col, quota2.col);
    
    // Much more restrictive - if they're too far apart, they're likely in different tables
    if (maxRow - minRow > 8 || maxCol - minCol > 8) {
      return false;
    }
    
    // Check if they're in the same row (horizontal table layout)
    if (quota1.row === quota2.row) {
      return true;
    }
    
    // For vertical proximity, check if there are connecting cells with borders
    let hasConnectingBorders = false;
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const cell = this.getCell(r, c);
        if (cell && (cell.borderStyles || cell.borders)) {
          hasConnectingBorders = true;
          break;
        }
      }
      if (hasConnectingBorders) break;
    }
    
    return hasConnectingBorders;
  }

  private isWorkContent(value: string): boolean {
    return value.includes('工作') && value.includes('内容') && value.includes('：');
  }

  private isNote(value: string): boolean {
    return value.startsWith('注') && (value.includes(':') || value.includes('：'));
  }

  private isContinuationTable(value: string): boolean {
    return value.includes('续表') || value.includes('（续）') || value.includes('(续)');
  }

  private parseChapterTitle(value: string): { number: string; name: string } | null {
    const match = value.match(/^第([一二三四五六七八九十\d]+)章\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        number: match[1],
        name: match[2].trim()
      };
    }
    return null;
  }

  private parseSectionTitle(value: string): { symbol: string; name: string } | null {
    // Match patterns like "第一节 减振装置安装"
    let match = value.match(/^(第[一二三四五六七八九十\d]+节)\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1],
        name: match[2].trim()
      };
    }
    
    // Match numbered patterns like "1.工地运输"
    match = value.match(/^(\d+)\.\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '.',
        name: match[2].trim()
      };
    }
    
    return null;
  }

  private parseSubSectionTitle(value: string): { symbol: string; name: string } | null {
    // Match patterns like "一 、减振装置安装" (with spaces)
    let match = value.match(/^([一二三四五六七八九十]+)\s+、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + ' 、',
        name: match[2].trim()
      };
    }
    
    // Match patterns like "一、减振装置安装" (without spaces)
    match = value.match(/^([一二三四五六七八九十]+)、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '、',
        name: match[2].trim()
      };
    }
    
    // Match patterns like "(1)单杆"
    match = value.match(/^\((\d+)\)\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: `(${match[1]})`,
        name: match[2].trim()
      };
    }
    
    return null;
  }

  private collectDescriptionText(startRow: number, endRow: number): string[] {
    const descriptions: string[] = [];
    
    for (let row = startRow; row <= endRow; row++) {
      const cellValue = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      
      // Only collect non-empty SimSun font cells (description text)
      if (cellValue && cell?.font?.name === 'SimSun') {
        // Skip if it looks like a structural element
        if (!this.isChapterTitle(cellValue, cell) && 
            !this.isSectionTitle(cellValue, cell) && 
            !this.isSubSectionTitle(cellValue, cell) &&
            !this.isNormCode(cellValue)) {
          descriptions.push(cellValue);
        }
      }
    }
    
    return descriptions;
  }

  private parseLeadingElements(startRow: number, endRow: number, startCol: number, endCol: number): TableStructure['leadingElements'] {
    // Look for work content and unit information in the leading rows
    for (let row = startRow; row <= Math.min(startRow + 3, endRow); row++) {
      for (let col = startCol; col <= endCol; col++) {
        const value = this.getCellValue(row, col);
        
        // Check for work content pattern
        if (this.isWorkContent(value)) {
          return {
            workContent: value,
            row: row
          };
        }
        
        // Check for unit pattern
        if (value && value.includes('单位') && value.includes('：')) {
          return {
            unit: value,
            row: row
          };
        }
      }
    }
    
    return undefined;
  }

  private parseNormCodesRow(startRow: number, endRow: number, startCol: number, endCol: number): TableStructure['normCodesRow'] {
    // Look for norm codes label cell (子目编号, 子目编码) - accounting for extra spaces
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= Math.min(startCol + 2, endCol); col++) {
        const value = this.getCellValue(row, col);
        const normalizedValue = value.replace(/\s+/g, ''); // Remove all spaces
        
        if (normalizedValue.includes('子目编号') || normalizedValue.includes('子目编码')) {
          // Found the label cell, now collect norm codes in the same row
          const normCodes: NormInfo[] = [];
          
          for (let normCol = col + 1; normCol <= endCol; normCol++) {
            const normValue = this.getCellValue(row, normCol);
            if (this.isNormCode(normValue)) {
              normCodes.push({
                code: normValue,
                row: row,
                col: normCol
              });
            }
          }
          
          if (normCodes.length > 0) {
            return {
              labelCell: value,
              normCodes: normCodes,
              row: row
            };
          }
        }
      }
    }
    
    return undefined;
  }

  private parseNormNamesRows(startRow: number, endRow: number, startCol: number, endCol: number, normCodesInfo?: Array<NormInfo>): TableStructure['normNamesRows'] {
    // Look for norm names label cell (子目名称) - accounting for extra spaces
    let labelRow = -1;
    let labelCell = '';
    
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= Math.min(startCol + 2, endCol); col++) {
        const value = this.getCellValue(row, col);
        const normalizedValue = value.replace(/\s+/g, ''); // Remove all spaces
        
        if (normalizedValue.includes('子目名称')) {
          labelRow = row;
          labelCell = value;
          break;
        }
      }
      if (labelRow > -1) break;
    }
    
    if (labelRow > -1 && normCodesInfo) {
      const normNames: Array<{
        baseName: string;
        spec?: string;
        unit?: string;
        fullName: string;
        normCode: string;
        col: number;
      }> = [];
      
      // Process each norm code column to get corresponding names and specs
      for (const normInfo of normCodesInfo) {
        const col = normInfo.col;
        const baseName = this.getCellValue(labelRow, col) || '';
        const spec = this.getCellValue(labelRow + 1, col) || '';
        
        // The unit for norm items is typically "台" or similar, not the consumption header
        // For most construction norm items, the default unit is "台" (set/unit)
        const unit = "台"; // Default unit for norm items
        
        // Form full name: ${baseName} ${spec}&${unit}
        let fullName = baseName;
        if (spec && spec !== baseName) {
          fullName += ` ${spec}`;
        }
        if (unit) {
          fullName += `&${unit}`;
        }
        
        normNames.push({
          baseName: baseName,
          spec: (spec && spec !== baseName) ? spec : undefined,
          unit: unit,
          fullName: fullName,
          normCode: normInfo.code,
          col: col
        });
      }
      
      return {
        labelCell: labelCell,
        normNames: normNames,
        startRow: labelRow,
        endRow: labelRow + 2 // Include potential unit row
      };
    }
    
    return undefined;
  }

  private parseResourcesSection(startRow: number, endRow: number, startCol: number, endCol: number, normCodes: NormInfo[]): TableStructure['resourcesSection'] {
    // Look for resources label cell (人材机名称, 单位, 消耗量) - accounting for extra spaces
    let resourcesStartRow = -1;
    let labelCell = '';
    let unitLabelCell = '';
    let consumptionLabelCell = '';
    
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= Math.min(startCol + 3, endCol); col++) {
        const value = this.getCellValue(row, col);
        const normalizedValue = value.replace(/\s+/g, ''); // Remove all spaces
        
        if (normalizedValue.includes('人材机名称') || normalizedValue === '名称') {
          resourcesStartRow = row;
          labelCell = value;
          
          // Look for unit and consumption labels in nearby cells
          for (let nearCol = col + 1; nearCol <= Math.min(col + 15, endCol); nearCol++) {
            const nearValue = this.getCellValue(row, nearCol);
            const normalizedNearValue = nearValue.replace(/\s+/g, '');
            if (normalizedNearValue.includes('单位')) {
              unitLabelCell = nearValue;
            }
            if (normalizedNearValue.includes('消耗量')) {
              consumptionLabelCell = nearValue;
            }
          }
          break;
        }
      }
      if (resourcesStartRow > -1) break;
    }
    
    if (resourcesStartRow > -1) {
      const resources: ResourceInfo[] = [];
      
      // Parse resource rows starting from the label row + 1
      let currentCategory = '';
      for (let row = resourcesStartRow + 1; row <= endRow; row++) {
        const categoryCell = this.getCellValue(row, startCol); // Column A - category
        const namesCell = this.getCellValue(row, startCol + 1); // Column B - names
        const unitsCell = this.getCellValue(row, startCol + 9); // Column J - units (typically column 10)
        
        if (!categoryCell && !namesCell) continue;
        
        // Skip label rows
        const normalizedCategoryCell = categoryCell.replace(/\s+/g, '');
        if (normalizedCategoryCell.includes('子目编号') || 
            normalizedCategoryCell.includes('子目名称') || 
            normalizedCategoryCell.includes('人材机名称')) {
          continue;
        }
        
        // Check if this is a category header (人工, 材料, 机械)
        if (normalizedCategoryCell === '人工' || categoryCell.includes('人工')) {
          currentCategory = '人工';
          
          // Parse resource names from column B
          if (namesCell) {
            const names = this.parseMultipleValues(namesCell);
            const units = this.parseMultipleValues(unitsCell || '');
            
            // Collect consumption data for each norm code
            const consumptionsArray: Array<{ [normCode: string]: number | string }> = [];
            
            for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
              const consumptions: { [normCode: string]: number | string } = {};
              
              for (const normInfo of normCodes) {
                const consumptionCell = this.getCellValue(row, normInfo.col);
                const consumptionValues = this.parseMultipleValues(consumptionCell || '');
                
                if (consumptionValues[nameIndex] && 
                    consumptionValues[nameIndex] !== '0' && 
                    consumptionValues[nameIndex] !== '-') {
                  const parsedConsumption = this.parseConsumptionValue(consumptionValues[nameIndex]);
                  // Store both the consumption value and primary flag info
                  consumptions[normInfo.code] = {
                    value: parsedConsumption.consumption,
                    originalString: consumptionValues[nameIndex],
                    isPrimary: parsedConsumption.isPrimary
                  };
                }
              }
              
              consumptionsArray.push(consumptions);
            }
            
            if (names.length > 0) {
              resources.push({
                category: currentCategory,
                names: names,
                units: units,
                consumptions: consumptionsArray,
                row: row
              });
            }
          }
        } else if (normalizedCategoryCell === '材料' || categoryCell.includes('材料')) {
          currentCategory = '材料';
          
          if (namesCell) {
            const names = this.parseMultipleValues(namesCell);
            const units = this.parseMultipleValues(unitsCell || '');
            
            const consumptionsArray: Array<{ [normCode: string]: number | string }> = [];
            
            for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
              const consumptions: { [normCode: string]: number | string } = {};
              
              for (const normInfo of normCodes) {
                const consumptionCell = this.getCellValue(row, normInfo.col);
                const consumptionValues = this.parseMultipleValues(consumptionCell || '');
                
                if (consumptionValues[nameIndex] && 
                    consumptionValues[nameIndex] !== '0' && 
                    consumptionValues[nameIndex] !== '-') {
                  const parsedConsumption = this.parseConsumptionValue(consumptionValues[nameIndex]);
                  // Store both the consumption value and primary flag info
                  consumptions[normInfo.code] = {
                    value: parsedConsumption.consumption,
                    originalString: consumptionValues[nameIndex],
                    isPrimary: parsedConsumption.isPrimary
                  };
                }
              }
              
              consumptionsArray.push(consumptions);
            }
            
            if (names.length > 0) {
              resources.push({
                category: currentCategory,
                names: names,
                units: units,
                consumptions: consumptionsArray,
                row: row
              });
            }
          }
        } else if (normalizedCategoryCell === '机械' || categoryCell.includes('机械')) {
          currentCategory = '机械';
          
          if (namesCell) {
            const names = this.parseMultipleValues(namesCell);
            const units = this.parseMultipleValues(unitsCell || '');
            
            const consumptionsArray: Array<{ [normCode: string]: number | string }> = [];
            
            for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
              const consumptions: { [normCode: string]: number | string } = {};
              
              for (const normInfo of normCodes) {
                const consumptionCell = this.getCellValue(row, normInfo.col);
                const consumptionValues = this.parseMultipleValues(consumptionCell || '');
                
                if (consumptionValues[nameIndex] && 
                    consumptionValues[nameIndex] !== '0' && 
                    consumptionValues[nameIndex] !== '-') {
                  const parsedConsumption = this.parseConsumptionValue(consumptionValues[nameIndex]);
                  // Store both the consumption value and primary flag info
                  consumptions[normInfo.code] = {
                    value: parsedConsumption.consumption,
                    originalString: consumptionValues[nameIndex],
                    isPrimary: parsedConsumption.isPrimary
                  };
                }
              }
              
              consumptionsArray.push(consumptions);
            }
            
            if (names.length > 0) {
              resources.push({
                category: currentCategory,
                names: names,
                units: units,
                consumptions: consumptionsArray,
                row: row
              });
            }
          }
        }
      }
      
      if (resources.length > 0) {
        return {
          labelCell: labelCell,
          unitLabelCell: unitLabelCell || undefined,
          consumptionLabelCell: consumptionLabelCell || undefined,
          resources: resources,
          startRow: resourcesStartRow,
          endRow: endRow
        };
      }
    }
    
    return undefined;
  }

  private buildNormResourceRelationships(normCodes: NormInfo[], resources: ResourceInfo[]): NormInfo[] {
    // Create a map to store category codes
    const categoryCodeMap: { [category: string]: number } = {
      '人工': 1,
      '材料': 2,
      '机械': 3,
    };

    // Build resource consumptions for each norm
    const updatedNorms: NormInfo[] = normCodes.map(norm => {
      const normResources: ResourceConsumption[] = [];
      
      for (const resource of resources) {
        const categoryCode = categoryCodeMap[resource.category] || 5;
        
        // Process each resource name in this category
        for (let nameIndex = 0; nameIndex < resource.names.length; nameIndex++) {
          const name = resource.names[nameIndex];
          const unit = resource.units[nameIndex] || '';
          const consumptionMap = resource.consumptions[nameIndex] || {};
          
          // Check if this norm has consumption for this resource
          if (consumptionMap[norm.code] !== undefined) {
            const consumptionValue = consumptionMap[norm.code];
            
            // Handle both old (number/string) and new (object) consumption formats
            let consumption: string;
            let isPrimary: boolean;
            
            if (typeof consumptionValue === 'object' && consumptionValue !== null && 'value' in consumptionValue) {
              // New format with consumption object
              consumption = String(consumptionValue.value);
              isPrimary = consumptionValue.isPrimary;
            } else {
              // Fallback: parse the consumption value directly
              const parsedConsumption = this.parseConsumptionValue(String(consumptionValue));
              consumption = parsedConsumption.consumption;
              isPrimary = parsedConsumption.isPrimary;
            }
            
            if (consumption !== '0' && consumption !== '-') {
              // If primary resource, set category code to 5
              const finalCategoryCode = isPrimary ? 5 : categoryCode;
              
              normResources.push({
                name: name,
                specification: '', // Will be filled from other sources if available
                unit: unit,
                consumption: consumption,
                isPrimary: isPrimary,
                category: resource.category,
                categoryCode: finalCategoryCode
              });
            }
          }
        }
      }
      
      return {
        ...norm,
        resources: normResources
      };
    });

    return updatedNorms;
  }

  // Helper method to parse multiple values from a single cell
  private parseMultipleValues(cellValue: string): string[] {
    if (!cellValue || cellValue.trim() === '') return [];
    
    // Split by common separators and clean up
    const values = cellValue.split(/[,，、\n\r]+/)
      .map(val => val.trim())
      .filter(val => val.length > 0);
      
    return values;
  }

  private parseConsumptionValue(value: string): { consumption: string; isPrimary: boolean } {
    if (!value || value.trim() === '' || value === '-' || value === '0') {
      return { consumption: '0', isPrimary: false };
    }
    
    const trimmedValue = value.trim();
    
    // Check if value is wrapped in parentheses (primary resource)
    const isPrimary = /^\(.*\)$/.test(trimmedValue);
    
    // Extract numeric value (remove parentheses if present)
    const numericValue = isPrimary 
      ? trimmedValue.slice(1, -1) // Remove parentheses
      : trimmedValue;
    
    // Keep the original string format to preserve trailing zeros
    return {
      consumption: numericValue,
      isPrimary
    };
  }

  private parseTrailingElements(startRow: number, endRow: number, startCol: number, endCol: number): TableStructure['trailingElements'] {
    const notes: string[] = [];
    const rows: number[] = [];
    
    // Look for notes in the trailing rows
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const value = this.getCellValue(row, col);
        if (this.isNote(value)) {
          notes.push(value);
          rows.push(row);
        }
      }
    }
    
    if (notes.length > 0) {
      return {
        notes: notes,
        rows: rows
      };
    }
    
    return undefined;
  }

  private detectTableAreas(): TableArea[] {
    const tableAreas: TableArea[] = [];
    
    // Strategy 1: Find all norm codes first, then build tables around them
    const normCodeCells: Array<{row: number; col: number; code: string}> = [];
    
    for (const cell of this.data.cells) {
      if (cell.value && typeof cell.value === 'string' && this.isNormCode(cell.value)) {
        normCodeCells.push({
          row: cell.row,
          col: cell.col,
          code: cell.value
        });
      }
    }
    
    console.log(`Found ${normCodeCells.length} norm code cells`);
    
    // Group norm codes by proximity (same table)
    const processedNorms = new Set<string>();
    
    for (const normCell of normCodeCells) {
      const key = `${normCell.row}-${normCell.col}`;
      if (processedNorms.has(key)) continue;
      
      // Find the table boundaries around this norm code
      const tableNorms: Array<{row: number; col: number; code: string}> = [];
      const visitedCells = new Set<string>();
      
      // Use BFS to find all connected norm codes in the same table structure
      const queue = [normCell];
      visitedCells.add(key);
      
      while (queue.length > 0) {
        const current = queue.shift()!;
        tableNorms.push(current);
        
        // Look for nearby norm codes within reasonable distance (much smaller radius)
        const searchRadius = 8; // cells - reduced for more granular tables
        for (const otherNorm of normCodeCells) {
          const otherKey = `${otherNorm.row}-${otherNorm.col}`;
          if (visitedCells.has(otherKey)) continue;
          
          // Check if this norm is within the same table area
          const rowDistance = Math.abs(otherNorm.row - current.row);
          const colDistance = Math.abs(otherNorm.col - current.col);
          
          // More restrictive conditions for smaller tables
          if (rowDistance <= searchRadius && colDistance <= searchRadius) {
            // Additional check: ensure they're in the same bordered area and close enough
            if (this.areInSameTable(current, otherNorm) && rowDistance <= 5) {
              visitedCells.add(otherKey);
              queue.push(otherNorm);
            }
          }
        }
      }
      
      // Mark all found norms as processed
      for (const tn of tableNorms) {
        processedNorms.add(`${tn.row}-${tn.col}`);
      }
      
      if (tableNorms.length > 0) {
        // Calculate table boundaries
        const rows = tableNorms.map(n => n.row);
        const cols = tableNorms.map(n => n.col);
        
        const minRow = Math.min(...rows);
        const maxRow = Math.max(...rows);
        const minCol = Math.min(...cols);
        const maxCol = Math.max(...cols);
        
        // Expand boundaries to include table structure
        const startRow = Math.max(1, minRow - 2);
        const endRow = Math.min(this.data.metadata.totalRows, maxRow + 10);
        const startCol = Math.max(1, minCol - 2);
        const endCol = Math.min(this.data.metadata.totalCols, maxCol + 5);
        
        // Extract table information
        const normCodes = tableNorms.map(n => n.code).sort();
        let unit = '';
        let workContent = '';
        const notes: string[] = [];
        
        // Look for work content in rows above the table
        for (let r = Math.max(1, startRow - 3); r < startRow + 3; r++) {
          for (let c = startCol; c <= endCol; c++) {
            const value = this.getCellValue(r, c);
            if (this.isWorkContent(value)) {
              workContent = value;
              break;
            }
            if (value && value.includes('单位') && value.includes('：')) {
              unit = value;
            }
          }
          if (workContent) break;
        }
        
        // Look for notes in rows below the table
        for (let r = maxRow; r <= Math.min(endRow + 5, this.data.metadata.totalRows); r++) {
          for (let c = startCol; c <= endCol; c++) {
            const value = this.getCellValue(r, c);
            if (this.isNote(value)) {
              notes.push(value);
            }
          }
        }
        
        // Parse detailed table structure - adjust range to include headers properly
        const tableStartRow = Math.max(1, minRow - 5); // Look a bit above the quota codes
        const tableEndRow = Math.min(this.data.metadata.totalRows, maxRow + 15); // Look a bit below
        
        const leadingElements = this.parseLeadingElements(tableStartRow, tableEndRow, 1, endCol);
        const normCodesRow = this.parseNormCodesRow(tableStartRow, tableEndRow, 1, endCol);
        const normNamesRows = this.parseNormNamesRows(tableStartRow, tableEndRow, 1, endCol, normCodesRow?.normCodes);
        const resourcesSection = normCodesRow ? this.parseResourcesSection(tableStartRow, tableEndRow, 1, endCol, normCodesRow.normCodes) : undefined;
        const trailingElements = this.parseTrailingElements(tableStartRow, tableEndRow, 1, endCol);

        const structure: TableStructure = {
          leadingElements,
          normCodesRow,
          normNamesRows,
          resourcesSection,
          trailingElements
        };

        // Build norm-resource relationships
        let norms: NormInfo[] | undefined = undefined;
        if (normCodesRow && resourcesSection) {
          norms = this.buildNormResourceRelationships(normCodesRow.normCodes, resourcesSection.resources);
        }

        const tableId = `table_${minRow}_${minCol}`;
        const table: TableArea = {
          id: tableId,
          range: { startRow, endRow, startCol, endCol },
          normCodes,
          unit,
          workContent: workContent || undefined, // Make optional
          notes,
          isContinuation: false,
          structure: structure,
          norms: norms
        };
        
        tableAreas.push(table);
        
        console.log(`  Found table at ${startRow}-${endRow}:${startCol}-${endCol} with ${normCodes.length} norms: ${normCodes.slice(0, 5).join(', ')}${normCodes.length > 5 ? '...' : ''}`);
        if (workContent) console.log(`    Work content: ${workContent.substring(0, 50)}...`);
        if (notes.length > 0) console.log(`    Notes: ${notes.length} found`);
        
        // Log detailed structure information
        if (structure.leadingElements) {
          console.log(`    Leading elements: ${structure.leadingElements.workContent ? 'work content' : ''}${structure.leadingElements.unit ? 'unit' : ''} at row ${structure.leadingElements.row}`);
        }
        if (structure.normCodesRow) {
          console.log(`    Norm codes row: ${structure.normCodesRow.normCodes.length} codes at row ${structure.normCodesRow.row}`);
        }
        if (structure.normNamesRows) {
          console.log(`    Norm names: ${structure.normNamesRows.normNames.length} norm names parsed`);
        }
        if (structure.resourcesSection) {
          console.log(`    Resources: ${structure.resourcesSection.resources.length} resources from row ${structure.resourcesSection.startRow}`);
        }
        if (structure.trailingElements) {
          console.log(`    Trailing elements: ${structure.trailingElements.notes.length} notes`);
        }
        if (norms && norms.length > 0) {
          const totalResources = norms.reduce((sum, norm) => sum + (norm.resources?.length || 0), 0);
          console.log(`    Norm-resource relationships: ${norms.length} norms with ${totalResources} total resource consumptions`);
        }
      }
    }

    // Process continuation tables
    this.processContinuationTables(tableAreas);

    return tableAreas;
  }

  private processContinuationTables(tableAreas: TableArea[]): void {
    for (let i = 0; i < tableAreas.length; i++) {
      const table = tableAreas[i];
      
      // Check if there's a "续表" indicator near this table
      for (let r = table.range.startRow - 2; r <= table.range.startRow + 2; r++) {
        const value = this.getCellValue(r, 1);
        if (this.isContinuationTable(value)) {
          table.isContinuation = true;
          
          // Find the previous table with the same quota codes
          for (let j = i - 1; j >= 0; j--) {
            const prevTable = tableAreas[j];
            if (prevTable.normCodes.some(code => table.normCodes.includes(code))) {
              table.continuationOf = prevTable.id;
              break;
            }
          }
          break;
        }
      }
    }
  }

  private buildHierarchicalStructure(tableAreas: TableArea[]): Chapter[] {
    const chapters: Chapter[] = [];
    let currentChapter: Chapter | null = null;
    let currentSection: Section | null = null;
    let currentSubSection: SubSection | null = null;

    // First pass: identify all structural headers
    const structuralHeaders: Array<{row: number, type: 'chapter' | 'section' | 'subsection', data: any}> = [];
    
    for (let row = 1; row <= this.data.metadata.totalRows; row++) {
      const cellValue = this.getCellValue(row, 1);
      if (!cellValue) continue;
      
      const cell = this.getCell(row, 1);

      if (this.isChapterTitle(cellValue, cell)) {
        const chapterInfo = this.parseChapterTitle(cellValue);
        if (chapterInfo) {
          structuralHeaders.push({row, type: 'chapter', data: chapterInfo});
        }
      } else if (this.isSectionTitle(cellValue, cell)) {
        const sectionInfo = this.parseSectionTitle(cellValue);
        if (sectionInfo) {
          structuralHeaders.push({row, type: 'section', data: sectionInfo});
        }
      } else if (this.isSubSectionTitle(cellValue, cell)) {
        const subSectionInfo = this.parseSubSectionTitle(cellValue);
        if (subSectionInfo) {
          structuralHeaders.push({row, type: 'subsection', data: subSectionInfo});
        }
      }
    }

    // Second pass: build hierarchy with descriptions
    for (let i = 0; i < structuralHeaders.length; i++) {
      const header = structuralHeaders[i];
      const nextHeader = structuralHeaders[i + 1];
      
      // Calculate description range (between current header and next header)
      const descriptionStart = header.row + 1;
      const descriptionEnd = nextHeader ? nextHeader.row - 1 : this.data.metadata.totalRows;
      
      if (header.type === 'chapter') {
        // Collect chapter description
        const description = this.collectDescriptionText(descriptionStart, Math.min(descriptionEnd, descriptionStart + 20));
        
        currentChapter = {
          id: `chapter_${header.data.number}`,
          name: header.data.name,
          number: header.data.number,
          description: description.length > 0 ? description : undefined,
          sections: [],
          tableAreas: []
        };
        chapters.push(currentChapter);
        currentSection = null;
        currentSubSection = null;
        this.processedRows.add(header.row);
        console.log(`Found chapter: ${header.data.name} (SimHei font confirmed)`);
        if (description.length > 0) {
          console.log(`  Chapter description: ${description.length} lines`);
        }
      } else if (header.type === 'section' && currentChapter) {
        // Collect section description
        const description = this.collectDescriptionText(descriptionStart, Math.min(descriptionEnd, descriptionStart + 15));
        
        currentSection = {
          id: `section_${currentChapter.id}_${header.data.symbol}`,
          name: header.data.name,
          number: header.data.symbol,
          description: description.length > 0 ? description : undefined,
          subSections: [],
          tableAreas: []
        };
        currentChapter.sections.push(currentSection);
        currentSubSection = null;
        this.processedRows.add(header.row);
        console.log(`Found section: ${header.data.name} (Font: SimHei)`);
        if (description.length > 0) {
          console.log(`  Section description: ${description.length} lines`);
        }
      } else if (header.type === 'subsection' && currentSection) {
        currentSubSection = {
          id: `subsection_${currentSection.id}_${header.data.symbol}`,
          name: header.data.name,
          level: 1,
          symbol: header.data.symbol,
          tableAreas: [],
          children: []
        };
        currentSection.subSections.push(currentSubSection);
        this.processedRows.add(header.row);
        console.log(`Found subsection: ${header.data.name} (Font: SimHei)`);
      }
    }

    // Assign table areas to appropriate hierarchy levels
    console.log(`Assigning ${tableAreas.length} table areas to hierarchy...`);
    
    for (const table of tableAreas) {
      const tableRow = table.range.startRow;
      let assigned = false;

      console.log(`Assigning table at row ${tableRow} with norms: ${table.normCodes.join(', ')}`);

      // Find the most recent chapter/section/subsection above this table by walking backwards
      let bestChapter: Chapter | null = null;
      let bestSection: Section | null = null;
      let bestSubSection: SubSection | null = null;
      
      let lastChapterRow = -1;
      let lastSectionRow = -1;
      let lastSubSectionRow = -1;

      // Walk backwards from the table to find the most recent hierarchy headers
      for (let row = tableRow - 1; row >= 1; row--) {
        const cellValue = this.getCellValue(row, 1);
        const cell = this.getCell(row, 1);
        
        if (!cellValue) continue;

        // Check for subsection (most specific)
        if (lastSubSectionRow === -1 && this.isSubSectionTitle(cellValue, cell)) {
          lastSubSectionRow = row;
          // Find the corresponding subsection in our hierarchy
          for (const chapter of chapters) {
            for (const section of chapter.sections) {
              for (const subSection of section.subSections) {
                if (cellValue.includes(subSection.name) || subSection.name.includes(cellValue.substring(0, 10))) {
                  bestSubSection = subSection;
                  bestSection = section;
                  bestChapter = chapter;
                  break;
                }
              }
              if (bestSubSection) break;
            }
            if (bestSubSection) break;
          }
        }
        
        // Check for section (medium specificity)
        if (lastSectionRow === -1 && this.isSectionTitle(cellValue, cell)) {
          lastSectionRow = row;
          // Only update if we haven't found a subsection
          if (!bestSubSection) {
            for (const chapter of chapters) {
              for (const section of chapter.sections) {
                if (cellValue.includes(section.name) || section.name.includes(cellValue.substring(0, 10))) {
                  bestSection = section;
                  bestChapter = chapter;
                  break;
                }
              }
              if (bestSection) break;
            }
          }
        }
        
        // Check for chapter (least specific)
        if (lastChapterRow === -1 && this.isChapterTitle(cellValue, cell)) {
          lastChapterRow = row;
          // Only update if we haven't found a section or subsection
          if (!bestSection && !bestSubSection) {
            for (const chapter of chapters) {
              if (cellValue.includes(chapter.name) || chapter.name.includes(cellValue.substring(0, 10))) {
                bestChapter = chapter;
                break;
              }
            }
          }
        }

        // If we've found all levels or gone too far back, stop searching
        if ((bestSubSection || bestSection || bestChapter) && row < tableRow - 50) {
          break;
        }
      }

      // Log what we found
      console.log(`  Found headers above table at row ${tableRow}:`);
      console.log(`    Chapter: ${bestChapter?.name || 'None'} (row ${lastChapterRow})`);
      console.log(`    Section: ${bestSection?.name || 'None'} (row ${lastSectionRow})`);
      console.log(`    SubSection: ${bestSubSection?.name || 'None'} (row ${lastSubSectionRow})`);

      // Assign to the most specific level found (prefer leaf nodes - subSections)
      if (bestSubSection) {
        bestSubSection.tableAreas.push(table);
        console.log(`  -> Assigned to subsection: ${bestSubSection.name}`);
        assigned = true;
      } else if (bestSection) {
        // If no subsection, assign to section (but prefer creating a default subsection)
        if (bestSection.subSections.length === 0) {
          // Create a default subsection for tables without explicit subsections
          const defaultSubSection: SubSection = {
            id: `subsection_${bestSection.id}_default`,
            name: 'Tables',
            level: 1,
            symbol: '',
            tableAreas: [table],
            children: []
          };
          bestSection.subSections.push(defaultSubSection);
          console.log(`  -> Created default subsection and assigned table to section: ${bestSection.name}`);
        } else {
          // Assign to the last subsection in this section
          const lastSubSection = bestSection.subSections[bestSection.subSections.length - 1];
          lastSubSection.tableAreas.push(table);
          console.log(`  -> Assigned to last subsection: ${lastSubSection.name} in section: ${bestSection.name}`);
        }
        assigned = true;
      } else if (bestChapter) {
        // If no section, assign to chapter (but prefer creating a default section/subsection)
        if (bestChapter.sections.length === 0) {
          // Create a default section and subsection
          const defaultSection: Section = {
            id: `section_${bestChapter.id}_default`,
            name: 'Default Section',
            number: '',
            subSections: [],
            tableAreas: []
          };
          
          const defaultSubSection: SubSection = {
            id: `subsection_${defaultSection.id}_default`,
            name: 'Tables',
            level: 1,
            symbol: '',
            tableAreas: [table],
            children: []
          };
          
          defaultSection.subSections.push(defaultSubSection);
          bestChapter.sections.push(defaultSection);
          console.log(`  -> Created default section/subsection and assigned to chapter: ${bestChapter.name}`);
        } else {
          // Assign to the last section's last subsection
          const lastSection = bestChapter.sections[bestChapter.sections.length - 1];
          if (lastSection.subSections.length === 0) {
            const defaultSubSection: SubSection = {
              id: `subsection_${lastSection.id}_default`,
              name: 'Tables',
              level: 1,
              symbol: '',
              tableAreas: [table],
              children: []
            };
            lastSection.subSections.push(defaultSubSection);
            console.log(`  -> Created default subsection in last section: ${lastSection.name}`);
          } else {
            const lastSubSection = lastSection.subSections[lastSection.subSections.length - 1];
            lastSubSection.tableAreas.push(table);
            console.log(`  -> Assigned to last subsection: ${lastSubSection.name}`);
          }
        }
        assigned = true;
      }

      if (!assigned) {
        console.log(`  -> WARNING: Could not assign table at row ${tableRow}`);
        // Create a fallback structure
        if (chapters.length > 0) {
          const firstChapter = chapters[0];
          if (firstChapter.sections.length === 0) {
            const defaultSection: Section = {
              id: `section_${firstChapter.id}_fallback`,
              name: 'Fallback Section',
              number: '',
              subSections: [],
              tableAreas: []
            };
            firstChapter.sections.push(defaultSection);
          }
          
          const lastSection = firstChapter.sections[firstChapter.sections.length - 1];
          if (lastSection.subSections.length === 0) {
            const defaultSubSection: SubSection = {
              id: `subsection_${lastSection.id}_fallback`,
              name: 'Fallback Tables',
              level: 1,
              symbol: '',
              tableAreas: [],
              children: []
            };
            lastSection.subSections.push(defaultSubSection);
          }
          
          const lastSubSection = lastSection.subSections[lastSection.subSections.length - 1];
          lastSubSection.tableAreas.push(table);
          console.log(`  -> Fallback: Assigned to subsection in first chapter`);
        }
      }
    }

    return chapters;
  }

  private findRowForText(text: string): number {
    for (let row = 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      if (value.includes(text)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextChapterRow(currentRow: number): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      if (this.isChapterTitle(value, cell)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextSectionRow(currentRow: number, chapter: Chapter): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      if (this.isSectionTitle(value, cell) || this.isChapterTitle(value, cell)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextSubSectionRow(currentRow: number, section: Section): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      if (this.isSubSectionTitle(value, cell) || this.isSectionTitle(value, cell) || this.isChapterTitle(value, cell)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  public parseStructure(): StructuredDocument {
    console.log('Starting structured parsing...');
    
    // Step 1: Detect all table areas
    console.log('Detecting table areas...');
    const tableAreas = this.detectTableAreas();
    console.log(`Found ${tableAreas.length} table areas`);

    // Step 2: Build hierarchical structure
    console.log('Building hierarchical structure...');
    const chapters = this.buildHierarchicalStructure(tableAreas);
    console.log(`Created ${chapters.length} chapters`);

    // Step 3: Create structured document
    const structuredDoc: StructuredDocument = {
      metadata: {
        ...this.data.metadata,
        structuredAt: new Date().toISOString()
      },
      chapters
    };

    console.log('Structured parsing completed');
    return structuredDoc;
  }
}

// Main execution
async function main() {
  const inputPath = './output/parsed-excel.json';
  const outputPath = './output/structured-excel.json';
  
  try {
    console.log(`Loading JSON data from: ${inputPath}`);
    const jsonData = JSON.parse(fs.readFileSync(inputPath, 'utf8')) as ParsedExcelData;
    
    const parser = new StructuredExcelParser(jsonData);
    const structuredDoc = parser.parseStructure();
    
    // Ensure output directory exists
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    // Write structured document
    fs.writeFileSync(outputPath, JSON.stringify(structuredDoc, null, 2), 'utf8');
    console.log(`Structured document saved to: ${outputPath}`);
    
    // Print summary
    console.log('\n=== PARSING SUMMARY ===');
    console.log(`Total chapters: ${structuredDoc.chapters.length}`);
    
    for (const chapter of structuredDoc.chapters) {
      console.log(`\n📖 Chapter ${chapter.number}: ${chapter.name}`);
      console.log(`   Sections: ${chapter.sections.length}, Tables: ${chapter.tableAreas.length}`);
      
      for (const section of chapter.sections) {
        console.log(`   📑 ${section.number} ${section.name}`);
        console.log(`      SubSections: ${section.subSections.length}, Tables: ${section.tableAreas.length}`);
        
        for (const subSection of section.subSections) {
          console.log(`      📄 ${subSection.symbol} ${subSection.name} (Tables: ${subSection.tableAreas.length})`);
        }
      }
    }
    
  } catch (error) {
    console.error('Error during structured parsing:', error);
    process.exit(1);
  }
}

// Run if called directly
if (require.main === module) {
  main();
}

export { StructuredExcelParser };