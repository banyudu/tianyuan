import * as ExcelJS from 'exceljs';

async function debugHierarchy() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('sample/input.xlsx');
    const worksheet = workbook.worksheets[0];

    console.log('=== 分析层级结构问题 ===');

    // Helper function to get cell value as string
    function getCellValue(row: number, col: number): string {
        const cell = worksheet.getCell(row, col);
        if (!cell.value) return '';
        
        if (typeof cell.value === 'object') {
            if ('richText' in cell.value && Array.isArray(cell.value.richText)) {
                return cell.value.richText.map((rt: any) => rt.text || '').join('');
            } else if ('text' in cell.value) {
                return String(cell.value.text);
            } else if ('result' in cell.value) {
                return String(cell.value.result);
            } else if ('formula' in cell.value) {
                return String(cell.value.result || cell.value.formula);
            }
        }
        
        return String(cell.value).trim();
    }

    // Find all subsections and their positions
    const subsections: Array<{row: number, text: string}> = [];
    const quotaCodes: Array<{row: number, code: string, col: number}> = [];
    
    for (let row = 1; row <= worksheet.rowCount; row++) {
        for (let col = 1; col <= 15; col++) {
            const value = getCellValue(row, col);
            
            // Find subsections (e.g., "三、VAV变风量空调机", "四、蓄冷(蓄热)设备")
            if (/^[一二三四五六七八九十]+、/.test(value)) {
                subsections.push({row, text: value});
            }
            
            // Find quota codes (1B-28, 1B-29, etc.)
            if (/^\d+[A-Z]-\d+$/.test(value)) {
                quotaCodes.push({row, code: value, col});
            }
        }
    }

    console.log('\n=== 找到的子节信息 ===');
    subsections.forEach(sub => {
        console.log(`行${sub.row}: ${sub.text}`);
    });

    console.log('\n=== 关键定额编号及其位置 ===');
    const targetCodes = ['1B-28', '1B-29', '1B-30', '1B-31', '1B-32', '1B-33', '1B-34', '1B-35', '1B-36'];
    const relevantCodes = quotaCodes.filter(q => targetCodes.includes(q.code));
    
    relevantCodes.forEach(q => {
        console.log(`行${q.row}, 列${q.col}: ${q.code}`);
    });

    console.log('\n=== 分析层级关系 ===');
    
    // Find the subsections around these codes
    const vavSubsection = subsections.find(s => s.text.includes('VAV变风量空调机'));
    const storageSubsection = subsections.find(s => s.text.includes('蓄冷') || s.text.includes('蓄热'));
    
    if (vavSubsection && storageSubsection) {
        console.log(`VAV变风量空调机 在行${vavSubsection.row}: ${vavSubsection.text}`);
        console.log(`蓄冷(蓄热)设备 在行${storageSubsection.row}: ${storageSubsection.text}`);
        
        console.log('\n问题分析:');
        console.log(`编号 1B-28 到 1B-36 的位置:`);
        relevantCodes.forEach(q => {
            if (q.row < storageSubsection.row) {
                console.log(`  ${q.code} (行${q.row}) - 错误：应该在 蓄冷(蓄热)设备 (行${storageSubsection.row}) 之后`);
            } else {
                console.log(`  ${q.code} (行${q.row}) - 正确：在 蓄冷(蓄热)设备 (行${storageSubsection.row}) 之后`);
            }
        });
    }

    // Look for sub-sub sections and sub-sub-sub sections  
    console.log('\n=== 查找子子节和子子子节 ===');
    for (let row = 200; row <= 300; row++) {
        const rowData: string[] = [];
        for (let col = 1; col <= 15; col++) {
            rowData.push(getCellValue(row, col));
        }
        const fullText = rowData.join(' ').trim();
        
        // Look for numbered subsections (1., 2., etc.) and parenthetical subsections ((1), (2), etc.)
        if (/^\d+\.|^\(\d+\)|^\([一二三四五六七八九十]+\)/.test(fullText)) {
            console.log(`行${row}: ${fullText.substring(0, 100)}...`);
        }
    }
}

debugHierarchy().catch(console.error);