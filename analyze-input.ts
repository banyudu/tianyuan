import * as ExcelJS from 'exceljs';

interface CellRange {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
    text: string;
}

function parseRange(range: string): CellRange | null {
    const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!match) return null;

    const [, startColStr, startRowStr, endColStr, endRowStr] = match;
    const startCol = columnToNumber(startColStr);
    const endCol = columnToNumber(endColStr);
    const startRow = parseInt(startRowStr);
    const endRow = parseInt(endRowStr);

    return { startRow, endRow, startCol, endCol, text: '' };
}

function columnToNumber(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
        result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
}

function numberToColumn(num: number): string {
    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

async function analyzeInputStructure(filePath: string): Promise<void> {
    console.log(`Analyzing input file: ${filePath}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
        console.log('No worksheet found');
        return;
    }

    console.log(`\nWorksheet: "${worksheet.name}"`);
    console.log(`Dimensions: ${worksheet.rowCount} rows, ${worksheet.columnCount} columns`);

    // Analyze merged cells to understand structure
    const mergedCells: CellRange[] = [];
    Object.keys(worksheet.model.merges || {}).forEach(range => {
        const parsed = parseRange(range);
        if (parsed) {
            // Get the text from the top-left cell of the merged range
            const cell = worksheet.getCell(parsed.startRow, parsed.startCol);
            parsed.text = cell.value?.toString() || '';
            mergedCells.push(parsed);
        }
    });

    console.log(`\nTotal merged cells: ${mergedCells.length}`);

    // Find header patterns
    console.log('\n=== HEADER ANALYSIS ===');
    for (let row = 1; row <= Math.min(20, worksheet.rowCount); row++) {
        const rowData: string[] = [];
        for (let col = 1; col <= worksheet.columnCount; col++) {
            const cell = worksheet.getCell(row, col);
            rowData.push(cell.value?.toString() || '');
        }

        // Only show rows with significant content
        const nonEmptyCount = rowData.filter(val => val.trim().length > 0).length;
        if (nonEmptyCount > 0) {
            console.log(`Row ${row}: [${nonEmptyCount} non-empty] ${rowData.slice(0, 10).join(' | ')}`);
        }
    }

    // Analyze column patterns
    console.log('\n=== COLUMN ANALYSIS ===');
    for (let col = 1; col <= Math.min(31, worksheet.columnCount); col++) {
        const colValues: string[] = [];
        for (let row = 1; row <= Math.min(50, worksheet.rowCount); row++) {
            const cell = worksheet.getCell(row, col);
            const value = cell.value?.toString() || '';
            if (value.trim()) {
                colValues.push(value);
            }
        }

        if (colValues.length > 0) {
            console.log(`Column ${numberToColumn(col)} (${col}): ${colValues.slice(0, 3).join(', ')}${colValues.length > 3 ? '...' : ''} (${colValues.length} values)`);
        }
    }

    // Look for data patterns
    console.log('\n=== DATA PATTERNS ===');
    const patterns = {
        codes: [] as string[],
        names: [] as string[],
        units: [] as string[],
        quantities: [] as string[]
    };

    for (let row = 1; row <= worksheet.rowCount; row++) {
        for (let col = 1; col <= worksheet.columnCount; col++) {
            const cell = worksheet.getCell(row, col);
            const value = cell.value?.toString() || '';

            if (value.trim()) {
                // Look for code patterns (numbers, alphanumeric)
                if (/^[A-Z0-9-]+$/.test(value) && value.length > 3) {
                    patterns.codes.push(value);
                }

                // Look for Chinese text (names/descriptions)
                if (/[\u4e00-\u9fff]/.test(value)) {
                    patterns.names.push(value);
                }

                // Look for units (m, kg, 个, etc.)
                if (/^(m|kg|个|台|套|t|L|cm|mm)$/.test(value)) {
                    patterns.units.push(value);
                }

                // Look for quantities (numbers)
                if (/^\d+(\.\d+)?$/.test(value)) {
                    patterns.quantities.push(value);
                }
            }
        }
    }

    console.log(`Found patterns:`);
    console.log(`- Codes: ${patterns.codes.slice(0, 5).join(', ')}... (${patterns.codes.length} total)`);
    console.log(`- Names: ${patterns.names.slice(0, 3).join(', ')}... (${patterns.names.length} total)`);
    console.log(`- Units: ${patterns.units.slice(0, 5).join(', ')}... (${patterns.units.length} total)`);
    console.log(`- Quantities: ${patterns.quantities.slice(0, 5).join(', ')}... (${patterns.quantities.length} total)`);
}

async function main() {
    const inputFile = 'data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx';
    await analyzeInputStructure(inputFile);
}

main().catch(console.error);
