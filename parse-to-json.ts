import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import { CellData, ParsedExcelData, BorderStyle } from './src/types';

class ExcelToJsonParser {
  private workbook: ExcelJS.Workbook;

  constructor() {
    this.workbook = new ExcelJS.Workbook();
  }

  private getCellValue(cell: ExcelJS.Cell): any {
    if (!cell || cell.value === null || cell.value === undefined) {
      return null;
    }

    if (typeof cell.value === 'object') {
      if ('richText' in cell.value && Array.isArray(cell.value.richText)) {
        return cell.value.richText.map((rt: any) => rt.text).join('');
      }
      if ('formula' in cell.value) {
        return cell.value.result || cell.value.formula;
      }
      if ('hyperlink' in cell.value) {
        return cell.value.text || cell.value.hyperlink;
      }
      if ('error' in cell.value) {
        return `#ERROR: ${cell.value.error}`;
      }
    }

    return cell.value;
  }

  private getCellType(cell: ExcelJS.Cell): string {
    if (!cell || cell.value === null || cell.value === undefined) {
      return 'empty';
    }

    if (typeof cell.value === 'object') {
      if ('richText' in cell.value) return 'richText';
      if ('formula' in cell.value) return 'formula';
      if ('hyperlink' in cell.value) return 'hyperlink';
      if ('error' in cell.value) return 'error';
    }

    if (typeof cell.value === 'string') return 'string';
    if (typeof cell.value === 'number') return 'number';
    if (typeof cell.value === 'boolean') return 'boolean';
    if (cell.value instanceof Date) return 'date';

    return 'unknown';
  }

  private getBorderInfo(cell: ExcelJS.Cell): CellData['borders'] {
    const borders = {
      top: false,
      bottom: false,
      left: false,
      right: false
    };

    if (cell.border) {
      borders.top = !!(cell.border.top && cell.border.top.style);
      borders.bottom = !!(cell.border.bottom && cell.border.bottom.style);
      borders.left = !!(cell.border.left && cell.border.left.style);
      borders.right = !!(cell.border.right && cell.border.right.style);
    }

    return borders;
  }

  private getBorderStyles(cell: ExcelJS.Cell): CellData['borderStyles'] {
    if (!cell.border) return undefined;

    const borderStyles: CellData['borderStyles'] = {};

    if (cell.border.top && cell.border.top.style) {
      borderStyles.top = {
        style: cell.border.top.style,
        color: this.getColorValue(cell.border.top.color)
      };
    }

    if (cell.border.bottom && cell.border.bottom.style) {
      borderStyles.bottom = {
        style: cell.border.bottom.style,
        color: this.getColorValue(cell.border.bottom.color)
      };
    }

    if (cell.border.left && cell.border.left.style) {
      borderStyles.left = {
        style: cell.border.left.style,
        color: this.getColorValue(cell.border.left.color)
      };
    }

    if (cell.border.right && cell.border.right.style) {
      borderStyles.right = {
        style: cell.border.right.style,
        color: this.getColorValue(cell.border.right.color)
      };
    }

    return Object.keys(borderStyles).length > 0 ? borderStyles : undefined;
  }

  private getColorValue(color: any): string | undefined {
    if (!color) return undefined;
    if (typeof color === 'string') return color;
    if (color.argb) return `#${color.argb.slice(2)}`; // Remove alpha channel
    if (color.rgb) return `#${color.rgb}`;
    if (color.theme !== undefined && color.tint !== undefined) {
      // For theme colors, return a description since we can't easily convert to hex
      return `theme:${color.theme},tint:${color.tint}`;
    }
    return undefined;
  }

  private getFillInfo(cell: ExcelJS.Cell): CellData['fill'] {
    if (!cell.fill) return undefined;

    const fill: CellData['fill'] = {};

    if (cell.fill.type) {
      fill.type = cell.fill.type;
    }

    if ('pattern' in cell.fill && cell.fill.pattern) {
      fill.pattern = cell.fill.pattern;
    }

    if ('fgColor' in cell.fill && cell.fill.fgColor) {
      fill.fgColor = this.getColorValue(cell.fill.fgColor);
    }

    if ('bgColor' in cell.fill && cell.fill.bgColor) {
      fill.bgColor = this.getColorValue(cell.fill.bgColor);
    }

    return Object.keys(fill).length > 0 ? fill : undefined;
  }

  private getFontInfo(cell: ExcelJS.Cell): CellData['font'] {
    if (!cell.font) return undefined;

    const font: CellData['font'] = {};

    if (cell.font.name) font.name = cell.font.name;
    if (cell.font.size) font.size = cell.font.size;
    if (cell.font.bold !== undefined) font.bold = cell.font.bold;
    if (cell.font.italic !== undefined) font.italic = cell.font.italic;
    if (cell.font.underline !== undefined) font.underline = !!cell.font.underline;
    if (cell.font.color) font.color = this.getColorValue(cell.font.color);

    return Object.keys(font).length > 0 ? font : undefined;
  }

  private getAlignmentInfo(cell: ExcelJS.Cell): CellData['alignment'] {
    if (!cell.alignment) return undefined;

    const alignment: CellData['alignment'] = {};

    if (cell.alignment.horizontal) alignment.horizontal = cell.alignment.horizontal;
    if (cell.alignment.vertical) alignment.vertical = cell.alignment.vertical;
    if (cell.alignment.wrapText !== undefined) alignment.wrapText = cell.alignment.wrapText;

    return Object.keys(alignment).length > 0 ? alignment : undefined;
  }

  private getMergedRangeInfo(worksheet: ExcelJS.Worksheet, row: number, col: number): CellData['mergedRange'] | undefined {
    const cell = worksheet.getCell(row, col);

    if (cell.isMerged && cell.master && cell.master.address === cell.address) {
      // This is the master cell, find its range
      const merges = (worksheet.model as any).merges || {};
      const cellAddress = cell.address;

      // Find the merge range that contains this master cell
      for (const rangeName of Object.keys(merges)) {
        const rangeString = merges[rangeName];
        if (typeof rangeString === 'string') {
          // Parse Excel range like "A1:AB1"
          const [startAddr, endAddr] = rangeString.split(':');
          const startCell = worksheet.getCell(startAddr);
          const endCell = worksheet.getCell(endAddr);

          if (startCell.address === cellAddress) {
            return {
              startRow: startCell.row as any as number,
              endRow: endCell.row as any as number,
              startCol: startCell.col as any as number,
              endCol: endCell.col as any as number
            };
          }
        }
      }
    }

    return undefined;
  }

  async parseExcelToJson(inputPath: string, outputPath: string): Promise<void> {
    console.log(`Loading Excel file: ${inputPath}`);

    await this.workbook.xlsx.readFile(inputPath);
    const worksheet = this.workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('No worksheet found in the Excel file');
    }

    console.log(`Sheet name: ${worksheet.name}`);
    console.log(`Dimensions: ${worksheet.rowCount} rows x ${worksheet.columnCount} columns`);

    const cells: CellData[] = [];
    const processedMergedRanges = new Set<string>();
    let actualRowCount = 0;
    let actualColCount = 0;

    // Process only rows and columns that actually contain data to improve performance
    const maxRow = Math.min(worksheet.rowCount, 1000);
    const maxCol = Math.min(worksheet.columnCount, 50);

    console.log(`Processing area: ${maxRow} rows x ${maxCol} columns`);

    // Iterate through rows and columns
    for (let row = 1; row <= maxRow; row++) {
      const worksheetRow = worksheet.getRow(row);

      for (let col = 1; col <= maxCol; col++) {
        const cell = worksheetRow.getCell(col);
        const value = this.getCellValue(cell);
        const cellType = this.getCellType(cell);

        // Track actual dimensions (non-empty cells)
        if (value !== null && value !== undefined && value !== '') {
          actualRowCount = Math.max(actualRowCount, row);
          actualColCount = Math.max(actualColCount, col);
        }

        // Check if this is a merged cell
        const isMerged = cell.isMerged;
        let mergedRange: CellData['mergedRange'] | undefined;

        // For merged cells, only process the master cell once
        if (isMerged) {
          const master = cell.master;
          if (master && master.address !== cell.address) {
            // This is not the master cell, skip it
            continue;
          }

          // This is the master cell, find the range
          mergedRange = this.getMergedRangeInfo(worksheet, row, col);
          const rangeKey = master ? master.address : cell.address;
          if (processedMergedRanges.has(rangeKey)) {
            continue;
          }
          processedMergedRanges.add(rangeKey);
        }

        // Only include cells that have content, borders, or are merged
        const hasBorders = this.getBorderInfo(cell);
        const hasContent = value !== null && value !== undefined && value !== '';

        if (hasContent || hasBorders.top || hasBorders.bottom || hasBorders.left || hasBorders.right || isMerged) {
          const cellData: CellData = {
            row,
            col,
            address: cell.address,
            value,
            type: cellType,
            merged: isMerged,
            mergedRange,
            borders: hasBorders
          };

          // Add style information
          const borderStyles = this.getBorderStyles(cell);
          if (borderStyles) cellData.borderStyles = borderStyles;

          const fill = this.getFillInfo(cell);
          if (fill) cellData.fill = fill;

          const font = this.getFontInfo(cell);
          if (font) cellData.font = font;

          const alignment = this.getAlignmentInfo(cell);
          if (alignment) cellData.alignment = alignment;

          cells.push(cellData);
        }
      }

      // Progress indicator
      if (row % 50 === 0) {
        console.log(`Processed ${row}/${maxRow} rows... (${cells.length} cells with data)`);
      }
    }

    const parsedData: ParsedExcelData = {
      metadata: {
        filename: path.basename(inputPath),
        sheetName: worksheet.name,
        totalRows: worksheet.rowCount,
        totalCols: worksheet.columnCount,
        actualRowCount,
        actualColCount,
        parsedAt: new Date().toISOString()
      },
      cells
    };

    console.log(`Writing JSON output to: ${outputPath}`);
    console.log(`Total cells processed: ${cells.length}`);
    console.log(`Actual data area: ${actualRowCount} rows x ${actualColCount} columns`);

    // Ensure output directory exists
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Write JSON file
    fs.writeFileSync(outputPath, JSON.stringify(parsedData, null, 2), 'utf8');

    console.log('Excel to JSON conversion completed successfully!');
  }
}

// Main execution
async function main() {
  const inputPath = './sample/input.xlsx';
  const outputPath = './output/parsed-excel.json';

  try {
    const parser = new ExcelToJsonParser();
    await parser.parseExcelToJson(inputPath, outputPath);
  } catch (error) {
    console.error('Error parsing Excel file:', error);
    process.exit(1);
  }
}

// Run if called directly
if (require.main === module) {
  main();
}

export { ExcelToJsonParser, CellData, ParsedExcelData };
