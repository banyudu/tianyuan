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

async function findDataSections(): Promise<void> {
    const inputFile = 'data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx';
    console.log(`Searching for data sections in: ${inputFile}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFile);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) return;

    console.log(`Total rows: ${worksheet.rowCount}, columns: ${worksheet.columnCount}\n`);

    // Look for potential data table headers
    const potentialHeaders = [
        '子目编号', '子目名称', '项目编号', '项目名称',
        '计量单位', '单位', '工作内容', '附注',
        '材料名称', '规格', '型号', '消耗量',
        '人工', '材料', '机械', '管理费'
    ];

    const dataSections: Array<{startRow: number, endRow: number, type: string, headers: string[]}> = [];

    // Scan through the entire file looking for data patterns
    for (let row = 1; row <= worksheet.rowCount; row++) {
        const rowValues: string[] = [];
        let hasTableHeaders = false;
        let headerCount = 0;

        // Get all values in this row
        for (let col = 1; col <= worksheet.columnCount; col++) {
            const cell = worksheet.getCell(row, col);
            const value = getCellValue(cell);
            rowValues.push(value);
        }

        // Check if this row contains table headers
        for (const value of rowValues) {
            if (potentialHeaders.some(header => value.includes(header))) {
                headerCount++;
                hasTableHeaders = true;
            }
        }

        if (hasTableHeaders && headerCount >= 3) {
            console.log(`\n=== POTENTIAL DATA SECTION AT ROW ${row} ===`);
            console.log(`Headers found: ${headerCount}`);
            console.log(`Row content: ${rowValues.filter(v => v.trim()).slice(0, 8).join(' | ')}`);

            // Look ahead to see how much data follows
            let dataRows = 0;
            for (let nextRow = row + 1; nextRow <= Math.min(row + 100, worksheet.rowCount); nextRow++) {
                let hasData = false;
                for (let col = 1; col <= Math.min(15, worksheet.columnCount); col++) {
                    const cell = worksheet.getCell(nextRow, col);
                    const value = getCellValue(cell);
                    if (value && value.trim() && !value.includes('·') && !value.includes('章')) {
                        hasData = true;
                        break;
                    }
                }
                if (hasData) dataRows++;
                else if (dataRows > 5) break; // Stop if we hit empty rows after finding data
            }

            console.log(`Data rows following: ${dataRows}`);

            if (dataRows > 5) {
                dataSections.push({
                    startRow: row,
                    endRow: row + dataRows,
                    type: 'data_table',
                    headers: rowValues.filter(v => v.trim())
                });

                // Show sample data
                console.log('\nSample data rows:');
                for (let sampleRow = row + 1; sampleRow <= Math.min(row + 5, worksheet.rowCount); sampleRow++) {
                    const sampleValues: string[] = [];
                    for (let col = 1; col <= Math.min(10, worksheet.columnCount); col++) {
                        const cell = worksheet.getCell(sampleRow, col);
                        const value = getCellValue(cell);
                        sampleValues.push(value || '');
                    }
                    if (sampleValues.some(v => v.trim())) {
                        console.log(`  Row ${sampleRow}: ${sampleValues.slice(0, 6).join(' | ')}`);
                    }
                }
            }
        }
    }

    console.log(`\n=== SUMMARY ===`);
    console.log(`Found ${dataSections.length} potential data sections:`);
    dataSections.forEach((section, index) => {
        console.log(`${index + 1}. Rows ${section.startRow}-${section.endRow} (${section.endRow - section.startRow + 1} rows)`);
        console.log(`   Headers: ${section.headers.slice(0, 5).join(', ')}...`);
    });

    // Look for specific patterns that might indicate the three types of tables we need
    console.log('\n=== SEARCHING FOR SPECIFIC TABLE TYPES ===');

        // Search for subitem info patterns
    const subitemPatterns = ['1B-', '2A-', '3C-', '子目', '编号'];
    let subitemSection: number | null = null;

    // Search for work content patterns
    const workContentPatterns = ['工作内容', '附注', '说明'];
    let workContentSection: number | null = null;

    // Search for material content patterns
    const materialPatterns = ['材料', '消耗量', '含量', '用量'];
    let materialSection: number | null = null;

    for (let row = 1; row <= worksheet.rowCount; row++) {
        const rowText: string[] = [];
        for (let col = 1; col <= Math.min(15, worksheet.columnCount); col++) {
            const cell = worksheet.getCell(row, col);
            const value = getCellValue(cell);
            rowText.push(value);
        }
        const rowContent = rowText.join(' ');

        // Check for subitem patterns
        if (subitemPatterns.some(pattern => rowContent.includes(pattern)) && !subitemSection) {
            const hasCode = /[0-9][A-Z]-[0-9]/.test(rowContent);
            if (hasCode) {
                subitemSection = row;
                console.log(`Potential subitem section at row ${row}: ${rowContent.slice(0, 100)}...`);
            }
        }

        // Check for work content patterns
        if (workContentPatterns.some(pattern => rowContent.includes(pattern)) && !workContentSection) {
            workContentSection = row;
            console.log(`Potential work content section at row ${row}: ${rowContent.slice(0, 100)}...`);
        }

        // Check for material patterns
        if (materialPatterns.some(pattern => rowContent.includes(pattern)) && !materialSection) {
            materialSection = row;
            console.log(`Potential material section at row ${row}: ${rowContent.slice(0, 100)}...`);
        }
    }
}

findDataSections().catch(console.error);
