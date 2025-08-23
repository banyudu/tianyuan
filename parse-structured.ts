import * as fs from 'fs';
import * as path from 'path';
import { CellData, Consumption, ParsedExcelData } from './src/types';
import {
  StructuredDocument,
  Chapter,
  Section,
  SubSection,
  TableArea,
  TableStructure,
  NormInfo,
  ResourceInfo,
  ResourceConsumption,
  BorderInfo
} from './src/structured-types';
import { ImprovedExcelConverter } from './convert';

const chopSpaces = (value: string) => value?.replace(/\s+/g, '')

const replaceParenthes = (value: string) => value?.replace(/（/g, '(').replace(/）/g, ')')

class StructuredExcelParser {
  private data: ParsedExcelData;
  private cellMap: Map<string, CellData>;
  private masterCellMap: Map<string, CellData>;
  private processedRows: Set<number>;
  private lastTableUnit: string = ''

  constructor(jsonData: ParsedExcelData) {
    this.data = jsonData;
    this.cellMap = new Map();
    this.masterCellMap = new Map()
    this.processedRows = new Set();

    // Create cell lookup map
    for (const cell of this.data.cells) {
      this.cellMap.set(`${cell.row}-${cell.col}`, cell);
      this.cellMap.set(cell.address, cell);

      // loop the merged area
      this.masterCellMap.set(`${cell.row}-${cell.col}`, cell);
      const mergedRange = cell.mergedRange
      if (mergedRange) {
        for (let i = mergedRange.startRow; i <= mergedRange.endRow; i++) {
          for (let j = mergedRange.startCol; j <= mergedRange.endCol; j++) {
            this.masterCellMap.set(`${i}-${j}`, cell);
          }
        }
      }
    }
  }

  private getCell(row: number, col: number): CellData | undefined {
    return this.cellMap.get(`${row}-${col}`);
  }

  private getMasterCell(row: number, col: number): CellData | undefined {
    return this.masterCellMap.get(`${row}-${col}`);
  }

  private getCellValue(row: number, col: number): string {
    const cell = this.getCell(row, col);
    if (!cell || !cell.value) return '';
    return String(cell.value).trim();
  }

  private getMasterCellValue(row: number, col: number): string {
    const cell = this.getMasterCell(row, col);
    if (!cell || !cell.value) return '';
    return String(cell.value).trim();
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
      return /^\(\d+\)/.test(value);
    }

    // Subsection titles with spaces like "一 、减振装置安装" use SimHei font
    const hasSpacedSubSectionPattern = /^[一二三四五六七八九十]+\s+、/.test(value);
    if (hasSpacedSubSectionPattern) {
      return true;
    }

    if (/^\s*\d+\./.test(value)) {
      return true
    }

    // Numbered subsections like "(1)"
    return /^\(\d+\)/.test(value);
  }

  private isNormCode(value: string): boolean {
    return /^\d+[A-Z]-\d+$/.test(value);
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
    const match = value.match(/^(第[一二三四五六七八九十\d]+章)\s*(.+?)(?:\s*·|$)/);
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

  private parseSubSectionTitle(value: string): { symbol: string; name: string; level: number } | null {
    // Match patterns like "一 、减振装置安装" (with spaces) - Level 1
    let match = value.match(/^([一二三四五六七八九十]+)\s+、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '、',
        name: match[2].trim(),
        level: 1
      };
    }

    // Match patterns like "一、减振装置安装" (without spaces) - Level 1
    match = value.match(/^([一二三四五六七八九十]+)、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '、',
        name: match[2].trim(),
        level: 1
      };
    }

    // Match numbered patterns like "1.工地运输", "2.底盘、拉盘、卡盘安装及电杆防腐" - Level 2
    match = value.match(/^(\d+)\.\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '.',
        name: match[2].trim(),
        level: 2
      };
    }

    // Match parenthetical patterns like "(1)单杆", "(2)接腿杆", "(3)撑杆及钢圈焊接" - Level 3
    match = value.match(/^\((\d+)\)\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: `(${match[1]})`,
        name: match[2].trim(),
        level: 3
      };
    }

    // Match circled number patterns like "①", "②" - Level 4
    match = value.match(/^([①②③④⑤⑥⑦⑧⑨⑩])\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1],
        name: match[2].trim(),
        level: 4
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
        unit?: string;
        fullName: string;
        normCode: string;
        col: number;
      }> = [];

      let tableUnit = "";

      // find the unit from the line above the cell of the first norm code
      const firstNormCell = normCodesInfo?.[0]

      const normNameLabelCell = this.getMasterCell(firstNormCell.row + 1, firstNormCell.col - 1)

      const normNameRowCount = (normNameLabelCell?.mergedRange?.endRow ?? 0) - (normNameLabelCell?.mergedRange?.startRow ?? 0) + 1

      if (firstNormCell) {
        const cellAboveTheFirstNormCode = this.getMasterCellValue(firstNormCell.row - 1, firstNormCell.col)
        // this cell may contains only the unit, or workcontent plus unit, or empty if this table is a continuious table of the last one.
        // let's find the unit using regexp
        const unitRegexp = /单\s*位\s*[：:]\s*([^\s]+)\s*/
        const match = String(cellAboveTheFirstNormCode).match(unitRegexp)
        if (match?.[1]) {
          tableUnit = match[1]
        } else {
          tableUnit = this.lastTableUnit
        }
      }

      const unitInTable = tableUnit === '见表' && normNameRowCount >= 2
      this.lastTableUnit = unitInTable ? '' : tableUnit

      // Process each norm code column to get corresponding names and specs
      for (const normInfo of normCodesInfo) {
        let normUnit = tableUnit
        const col = normInfo.col;
        const names = []
        const usedAddrs = new Set()
        for (let i = 0; i < normNameRowCount; i++) {
          const masterCell = this.getMasterCell(labelRow + i, col);
          if (usedAddrs.has(masterCell?.address)) {
            continue;
          }
          usedAddrs.add(masterCell?.address)
          const name = this.getMasterCellValue(labelRow + i, col) || '';
          names.push(name);
        }

        if (unitInTable) {
          normUnit = names.pop() as string;
        }

        const baseName = replaceParenthes((names.map(item => item.replace(/\s+/g, '')).join(' ')))

        // Form full name: ${baseName} ${specUnit} ${spec}&${unit}
        const fullName = replaceParenthes(`${baseName}&${normUnit}`)

        normNames.push({
          baseName: baseName,
          unit: normUnit,
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
            const consumptionsArray: Array<{ [normCode: string]: Consumption }> = [];

            for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
              const consumptions: { [normCode: string]: Consumption } = {};

              for (const normInfo of normCodes) {
                const consumptionCell = this.getCellValue(row, normInfo.col);
                const consumptionValues = this.parseMultipleValues(consumptionCell || '');

                if (consumptionValues[nameIndex] &&
                  consumptionValues[nameIndex] !== '0' &&
                  consumptionValues[nameIndex] !== '-') {

                  const parsedConsumption = this.parseConsumptionValue(consumptionValues[nameIndex]);
                  // Store both the consumption value and primary flag info
                  consumptions[normInfo.code] = {
                    value: parsedConsumption.value,
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

            const consumptionsArray: Array<Record<string, Consumption>> = [];

            for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
              const consumptions: Record<string, Consumption> = {}

              for (const normInfo of normCodes) {
                const consumptionCell = this.getCellValue(row, normInfo.col);
                const consumptionValues = this.parseMultipleValues(consumptionCell || '');

                if (consumptionValues[nameIndex] &&
                  consumptionValues[nameIndex] !== '0' &&
                  consumptionValues[nameIndex] !== '-') {
                  const parsedConsumption = this.parseConsumptionValue(consumptionValues[nameIndex]);
                  // Store both the consumption value and primary flag info
                  consumptions[normInfo.code] = parsedConsumption
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

            const consumptionsArray: Array<Record<string, Consumption>> = [];

            for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
              const consumptions: Record<string, Consumption> = {};

              for (const normInfo of normCodes) {
                const consumptionCell = this.getCellValue(row, normInfo.col);
                const consumptionValues = this.parseMultipleValues(consumptionCell || '');

                if (consumptionValues[nameIndex] &&
                  consumptionValues[nameIndex] !== '0' &&
                  consumptionValues[nameIndex] !== '-') {
                  const parsedConsumption = this.parseConsumptionValue(consumptionValues[nameIndex]);
                  // Store both the consumption value and primary flag info
                  consumptions[normInfo.code] = parsedConsumption
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
            let consumption: string | null;
            let isPrimary: boolean;

            if (typeof consumptionValue === 'object' && consumptionValue !== null && 'value' in consumptionValue) {
              // New format with consumption object
              consumption = consumptionValue.value;
              isPrimary = consumptionValue.isPrimary;
            } else {
              // Fallback: parse the consumption value directly
              const parsedConsumption = this.parseConsumptionValue(String(consumptionValue));
              consumption = parsedConsumption.value;
              isPrimary = !!parsedConsumption.isPrimary;
            }

            if (consumption !== '0' && String(consumption) !== '-') {
              // If primary resource, set category code to 5
              const finalCategoryCode = isPrimary ? 5 : categoryCode;

              normResources.push({
                name: name,
                specification: '', // Will be filled from other sources if available
                unit: unit,
                consumption,
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

  private parseConsumptionValue(value: string): Consumption {
    if (!value || value.trim() === '' || value === '-' || value === '0') {
      return { value: '0', isPrimary: false, originalString: value };
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
      value: numericValue,
      isPrimary,
      originalString: trimmedValue
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
    const normCodeCells: Array<{ row: number; col: number; code: string }> = [];

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
      const tableNorms: Array<{ row: number; col: number; code: string }> = [];
      const visitedCells = new Set<string>();

      // Use BFS to find all connected norm codes in the same table structure
      const queue = [normCell];
      visitedCells.add(key);

      while (queue.length > 0) {
        const current = queue.shift()!;
        tableNorms.push(current);

        // Look for nearby norm codes within reasonable distance (much smaller radius)
        for (const otherNorm of normCodeCells) {
          const otherKey = `${otherNorm.row}-${otherNorm.col}`;
          if (visitedCells.has(otherKey)) continue;

          // Check if this norm is within the same table area
          const rowDistance = Math.abs(otherNorm.row - current.row);

          // More restrictive conditions for smaller tables
          if (rowDistance === 0) {
            // Additional check: ensure they're in the same bordered area and close enough
            visitedCells.add(otherKey);
            queue.push(otherNorm);
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
            if (value && value.includes('单位') && value.includes('：')) {
              unit = value;
            }
            if (this.isWorkContent(value)) {
              workContent = value;
              break;
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

  private findParentSubSection(section: Section, targetLevel: number): SubSection | null {
    // Find the most recent subsection at the target level
    const findInSubSections = (subSections: SubSection[]): SubSection | null => {
      for (let i = subSections.length - 1; i >= 0; i--) {
        const sub = subSections[i];
        if (sub.level === targetLevel) {
          return sub;
        }
        // Recursively search in children
        const found = findInSubSections(sub.children);
        if (found) return found;
      }
      return null;
    };

    return findInSubSections(section.subSections);
  }

  private buildHierarchicalStructure(tableAreas: TableArea[]): Chapter[] {
    const chapters: Chapter[] = [];

    // Step 1: Detect all structural headers with their row numbers
    const structuralHeaders = this.detectAllStructuralHeaders();

    // Step 2: Build hierarchical structure
    this.buildHierarchy(structuralHeaders, chapters);

    // Step 3: Assign tables to nearest sections
    this.assignTablesToSections(tableAreas, structuralHeaders, chapters);

    return chapters;
  }

  private detectAllStructuralHeaders(): Array<{ row: number, type: 'chapter' | 'section' | 'subsection', data: any }> {
    const headers: Array<{ row: number, type: 'chapter' | 'section' | 'subsection', data: any }> = [];

    for (let row = 1; row <= this.data.metadata.totalRows; row++) {
      for (let col = 1; col <= Math.min(15, this.data.metadata.totalCols); col++) {
        const cellValue = this.getCellValue(row, col);
        if (!cellValue) continue;

        const cell = this.getCell(row, col);

        // Chapters and sections are typically in column 1
        if (col === 1) {
          if (this.isChapterTitle(cellValue, cell)) {
            const chapterInfo = this.parseChapterTitle(cellValue);
            if (chapterInfo) {
              headers.push({ row, type: 'chapter', data: chapterInfo });
              console.log(`Found chapter at row ${row}: ${chapterInfo.name}`);
            }
          } else if (this.isSectionTitle(cellValue, cell)) {
            const sectionInfo = this.parseSectionTitle(cellValue);
            if (sectionInfo) {
              headers.push({ row, type: 'section', data: sectionInfo });
              console.log(`Found section at row ${row}: ${sectionInfo.name}`);
            }
          }
        }

        // Subsections can be in any column
        if (this.isSubSectionTitle(cellValue, cell)) {
          const subSectionInfo = this.parseSubSectionTitle(cellValue);
          if (subSectionInfo) {
            headers.push({ row, type: 'subsection', data: { ...subSectionInfo, col } });
            console.log(`Found subsection at row ${row}: ${subSectionInfo.name} (level ${subSectionInfo.level})`);
          }
        }
      }
    }

    return headers.sort((a, b) => a.row - b.row);
  }

  private buildHierarchy(structuralHeaders: Array<{ row: number, type: 'chapter' | 'section' | 'subsection', data: any }>, chapters: Chapter[]): void {
    let currentChapter: Chapter | null = null;
    let currentSection: Section | null = null;

    for (const header of structuralHeaders) {
      if (header.type === 'chapter') {
        currentChapter = {
          id: `chapter_${header.data.number}`,
          name: replaceParenthes(header.data.name).replace(/\s+/g, ''),
          number: header.data.number,
          row: header.row,
          sections: [],
          tableAreas: []
        };
        chapters.push(currentChapter);
        currentSection = null;
        this.processedRows.add(header.row);
      } else if (header.type === 'section' && currentChapter) {
        currentSection = {
          id: `section_${currentChapter.id}_${header.data.symbol}`,
          name: replaceParenthes(header.data.name).replace(/\s+/g, ''),
          number: header.data.symbol,
          row: header.row,
          subSections: [],
          tableAreas: []
        };
        currentChapter.sections.push(currentSection);
        this.processedRows.add(header.row);
      } else if (header.type === 'subsection' && currentSection) {
        const subsectionData = header.data;
        const newSubSection: SubSection = {
          id: `subsection_${currentSection.id}_${subsectionData.symbol}`,
          name: replaceParenthes(subsectionData.name).replace(/\s+/g, ''),
          level: subsectionData.level,
          symbol: subsectionData.symbol,
          row: header.row,
          tableAreas: [],
          children: []
        };

        // Handle multi-level hierarchy
        if (subsectionData.level === 1) {
          currentSection.subSections.push(newSubSection);
        } else if (subsectionData.level > 1) {
          const parentSubSection = this.findParentSubSection(currentSection, subsectionData.level - 1);
          if (parentSubSection) {
            newSubSection.parentId = parentSubSection.id;
            parentSubSection.children.push(newSubSection);
          } else {
            currentSection.subSections.push(newSubSection);
          }
        }

        this.processedRows.add(header.row);
      }
    }
  }

  private assignTablesToSections(tableAreas: TableArea[], structuralHeaders: Array<{ row: number, type: 'chapter' | 'section' | 'subsection', data: any }>, chapters: Chapter[]): void {
    console.log(`Assigning ${tableAreas.length} table areas to hierarchy...`);

    for (const table of tableAreas) {
      // Use the norm code row instead of table startRow for more accurate assignment
      const normCodeRow = table.structure?.normCodesRow?.row || table.range.startRow;
      console.log(`Assigning table with norm codes at row ${normCodeRow}, norms: ${table.normCodes.join(', ')}`);

      // Find the nearest section/subsection above the norm code row
      this.assignTableToNearestHierarchyLevel(table, normCodeRow, chapters);
    }
  }

  private assignTableToNearestHierarchyLevel(table: TableArea, normCodeRow: number, chapters: Chapter[]): void {
    let bestChapter: Chapter | null = null;
    let bestSection: Section | null = null;
    let bestSubSection: SubSection | null = null;

    let minChapterDistance = Infinity;
    let minSectionDistance = Infinity;
    let minSubSectionDistance = Infinity;

    // Find the nearest chapter, section, and subsection above the norm code row
    for (const chapter of chapters) {
      if (chapter.row < normCodeRow) {
        const distance = normCodeRow - chapter.row;
        if (distance < minChapterDistance) {
          minChapterDistance = distance;
          bestChapter = chapter;
        }
      }

      for (const section of chapter.sections) {
        if (section.row < normCodeRow) {
          const distance = normCodeRow - section.row;
          if (distance < minSectionDistance) {
            minSectionDistance = distance;
            bestSection = section;
          }
        }

        // Search in subsections recursively
        const foundSubSection = this.findNearestSubSection(section.subSections, normCodeRow);
        if (foundSubSection) {
          const distance = normCodeRow - foundSubSection.row;
          if (distance < minSubSectionDistance) {
            minSubSectionDistance = distance;
            bestSubSection = foundSubSection;
          }
        }
      }
    }

    // Assign to the most specific (nearest) level found
    if (bestSubSection) {
      bestSubSection.tableAreas.push(table);
      console.log(`  -> Assigned to subsection: ${bestSubSection.name} (distance: ${minSubSectionDistance})`);
    } else if (bestSection) {
      // If no subsection found, create a fallback subsection or assign to the section's last subsection
      if (bestSection.subSections.length > 0) {
        const lastSubSection = bestSection.subSections[bestSection.subSections.length - 1];
        lastSubSection.tableAreas.push(table);
        console.log(`  -> Assigned to section's last subsection: ${lastSubSection.name} (section distance: ${minSectionDistance})`);
      } else {
        bestSection.tableAreas?.push(table);
        console.log(`  -> Assigned to section: ${bestSection.name} (distance: ${minSectionDistance})`);
      }
    } else if (bestChapter) {
      // Fallback to chapter level
      if (bestChapter.sections.length > 0) {
        const lastSection = bestChapter.sections[bestChapter.sections.length - 1];
        if (lastSection.subSections.length > 0) {
          const lastSubSection = lastSection.subSections[lastSection.subSections.length - 1];
          lastSubSection.tableAreas.push(table);
          console.log(`  -> Assigned to chapter's last subsection: ${lastSubSection.name} (chapter distance: ${minChapterDistance})`);
        } else {
          lastSection.tableAreas?.push(table);
          console.log(`  -> Assigned to chapter's last section: ${lastSection.name} (chapter distance: ${minChapterDistance})`);
        }
      } else {
        bestChapter.tableAreas?.push(table);
        console.log(`  -> Assigned to chapter: ${bestChapter.name} (distance: ${minChapterDistance})`);
      }
    } else {
      console.log(`  -> WARNING: Could not find any header above norm code row ${normCodeRow}`);
      this.assignToFallback(table, chapters);
    }
  }

  private findNearestSubSection(subSections: SubSection[], normCodeRow: number): SubSection | null {
    let nearest: SubSection | null = null;
    let minDistance = Infinity;

    for (const subSection of subSections) {
      if (subSection.row < normCodeRow) {
        const distance = normCodeRow - subSection.row;
        if (distance < minDistance) {
          minDistance = distance;
          nearest = subSection;
        }
      }

      // Search in children recursively
      const childResult = this.findNearestSubSection(subSection.children, normCodeRow);
      if (childResult && childResult.row < normCodeRow) {
        const distance = normCodeRow - childResult.row;
        if (distance < minDistance) {
          minDistance = distance;
          nearest = childResult;
        }
      }
    }

    return nearest;
  }


  private assignToFallback(table: TableArea, chapters: Chapter[]): void {
    if (chapters.length > 0) {
      const firstChapter = chapters[0];
      if (firstChapter.sections.length > 0) {
        const lastSection = firstChapter.sections[firstChapter.sections.length - 1];
        if (lastSection.subSections.length > 0) {
          const lastSubSection = lastSection.subSections[lastSection.subSections.length - 1];
          lastSubSection.tableAreas.push(table);
          console.log(`  -> Fallback: Assigned to last subsection`);
        } else {
          lastSection.tableAreas?.push(table);
          console.log(`  -> Fallback: Assigned to last section`);
        }
      } else {
        firstChapter.tableAreas?.push(table);
        console.log(`  -> Fallback: Assigned to first chapter`);
      }
    }
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

  // New method to generate improved CSV outputs using the improved converter
  public async generateImprovedOutputs(inputFilePath: string, outputDir: string = './output'): Promise<void> {
    console.log('\n=== 生成改进版CSV输出 ===');

    try {
      const converter = new ImprovedExcelConverter();
      await converter.loadFile(inputFilePath);

      // Export to the specified output directory
      const improvedOutputDir = path.join(outputDir, 'csv');
      await converter.exportToCsv(improvedOutputDir);

      console.log(`改进版CSV文件已生成到: ${improvedOutputDir}`);
    } catch (error) {
      console.error('生成改进版CSV输出时发生错误:', error);
      throw error;
    }
  }
}

// Main execution
async function main() {
  const inputPath = './output/parsed-excel.json';
  const outputPath = './output/structured-excel.json';
  const inputExcelPath = './sample/input.xlsx'; // Path to original Excel file

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

    // Generate improved CSV outputs with fixed hierarchical structure
    if (fs.existsSync(inputExcelPath)) {
      await parser.generateImprovedOutputs(inputExcelPath, outputDir);
    } else {
      console.log(`警告: 原始Excel文件未找到 (${inputExcelPath})，跳过改进版CSV生成`);
    }

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
