import * as ExcelJS from 'exceljs';

function getCellValue(cell: ExcelJS.Cell): string {
    if (!cell.value) return '';

    // Handle rich text
    if (typeof cell.value === 'object' && 'richText' in cell.value) {
        return cell.value.richText.map((rt: any) => rt.text).join('');
    }

    // Handle formulas
    if (typeof cell.value === 'object' && 'formula' in cell.value) {
        return cell.value.result?.toString() || '';
    }

    // Handle shared strings
    if (typeof cell.value === 'object' && 'sharedString' in cell.value) {
        return (cell.value as any).sharedString.toString();
    }

    return cell.value.toString();
}

async function detailedAnalysis(): Promise<void> {
    const inputFile = 'data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx';
    console.log(`Detailed analysis of: ${inputFile}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFile);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) return;

    console.log(`Worksheet: "${worksheet.name}"`);
    console.log(`Dimensions: ${worksheet.rowCount} rows, ${worksheet.columnCount} columns\n`);

    // Find actual table start by looking for meaningful headers
    let tableStartRow = 1;
    let headerRow = 1;

    console.log('=== FINDING TABLE STRUCTURE ===');

    // Look for the actual data table (skip title/header sections)
    for (let row = 1; row <= Math.min(50, worksheet.rowCount); row++) {
        const rowValues: string[] = [];
        let meaningfulCells = 0;

        for (let col = 1; col <= Math.min(10, worksheet.columnCount); col++) {
            const cell = worksheet.getCell(row, col);
            const value = getCellValue(cell);
            rowValues.push(value);

            if (value && value.trim() && !value.includes('建设工程消耗量标准') && !value.includes('补充子目') && !value.includes('目  录')) {
                meaningfulCells++;
            }
        }

        if (meaningfulCells >= 3) {
            console.log(`Row ${row} (${meaningfulCells} meaningful): ${rowValues.slice(0, 5).join(' | ')}`);

            // Look for header patterns
            const rowText = rowValues.join(' ').toLowerCase();
            if (rowText.includes('子目编号') || rowText.includes('子目名称') || rowText.includes('编号') || rowText.includes('名称')) {
                headerRow = row;
                tableStartRow = row + 1;
                console.log(`*** Potential header row found: ${row} ***`);
            }
        }
    }

    console.log(`\nUsing header row: ${headerRow}, data starts at: ${tableStartRow}\n`);

    // Analyze the header structure
    console.log('=== HEADER STRUCTURE ===');
    const headerValues: string[] = [];
    for (let col = 1; col <= worksheet.columnCount; col++) {
        const cell = worksheet.getCell(headerRow, col);
        const value = getCellValue(cell);
        headerValues.push(value);
        if (value && value.trim()) {
            console.log(`Column ${col} (${String.fromCharCode(64 + col)}): "${value}"`);
        }
    }

    // Sample data rows
    console.log('\n=== SAMPLE DATA ROWS ===');
    for (let row = tableStartRow; row <= Math.min(tableStartRow + 10, worksheet.rowCount); row++) {
        const rowValues: string[] = [];
        let hasData = false;

        for (let col = 1; col <= Math.min(15, worksheet.columnCount); col++) {
            const cell = worksheet.getCell(row, col);
            const value = getCellValue(cell);
            rowValues.push(value);
            if (value && value.trim()) hasData = true;
        }

        if (hasData) {
            console.log(`Row ${row}: ${rowValues.slice(0, 8).join(' | ')}`);
        }
    }

    // Look for patterns in the data
    console.log('\n=== DATA PATTERN ANALYSIS ===');
    const patterns = {
        codes: new Set<string>(),
        names: new Set<string>(),
        units: new Set<string>(),
        numbers: new Set<string>()
    };

    for (let row = tableStartRow; row <= worksheet.rowCount; row++) {
        for (let col = 1; col <= worksheet.columnCount; col++) {
            const cell = worksheet.getCell(row, col);
            const value = getCellValue(cell);

            if (value && value.trim()) {
                // Item codes (alphanumeric patterns)
                if (/^[A-Z0-9][A-Z0-9\-\.]*$/.test(value) && value.length > 1) {
                    patterns.codes.add(value);
                }

                // Chinese text (descriptions)
                if (/[\u4e00-\u9fff]/.test(value) && value.length > 2) {
                    patterns.names.add(value);
                }

                // Units
                if (/^(m|kg|个|台|套|t|L|cm|mm|m²|m³|只|根|块|张)$/.test(value)) {
                    patterns.units.add(value);
                }

                // Numbers
                if (/^\d+(\.\d+)?$/.test(value)) {
                    patterns.numbers.add(value);
                }
            }
        }
    }

    console.log(`Unique codes: ${Array.from(patterns.codes).slice(0, 10).join(', ')}... (${patterns.codes.size} total)`);
    console.log(`Units found: ${Array.from(patterns.units).join(', ')}`);
    console.log(`Sample names: ${Array.from(patterns.names).slice(0, 3).join(', ')}...`);
    console.log(`Numbers range: ${Array.from(patterns.numbers).slice(0, 10).join(', ')}... (${patterns.numbers.size} total)`);

    // Identify column purposes based on content patterns
    console.log('\n=== COLUMN PURPOSE ANALYSIS ===');
    for (let col = 1; col <= Math.min(20, worksheet.columnCount); col++) {
        const columnValues: string[] = [];
        const columnPatterns = { codes: 0, names: 0, units: 0, numbers: 0 };

        for (let row = tableStartRow; row <= Math.min(tableStartRow + 100, worksheet.rowCount); row++) {
            const cell = worksheet.getCell(row, col);
            const value = getCellValue(cell);

            if (value && value.trim()) {
                columnValues.push(value);

                if (/^[A-Z0-9][A-Z0-9\-\.]*$/.test(value)) columnPatterns.codes++;
                if (/[\u4e00-\u9fff]/.test(value)) columnPatterns.names++;
                if (/^(m|kg|个|台|套|t|L|cm|mm|m²|m³|只|根|块|张)$/.test(value)) columnPatterns.units++;
                if (/^\d+(\.\d+)?$/.test(value)) columnPatterns.numbers++;
            }
        }

        if (columnValues.length > 0) {
            const purpose = Object.entries(columnPatterns).reduce((a, b) => columnPatterns[a[0] as keyof typeof columnPatterns] > columnPatterns[b[0] as keyof typeof columnPatterns] ? a : b)[0];
            console.log(`Column ${String.fromCharCode(64 + col)} (${col}): ${purpose} (${columnPatterns[purpose as keyof typeof columnPatterns]}/${columnValues.length}) - "${columnValues[0]}"`);
        }
    }
}

detailedAnalysis().catch(console.error);
