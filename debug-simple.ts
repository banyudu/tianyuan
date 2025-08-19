import * as ExcelJS from 'exceljs';

async function analyzeStructure() {
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

    // Find specific subsections
    const subsections: Array<{row: number, text: string}> = [];
    const quotaCodes: Array<{row: number, code: string}> = [];
    
    // Focus on actual content area (skip table of contents)
    for (let row = 60; row <= 250; row++) {
        const firstCol = getCellValue(row, 1);
        
        // Find subsections (e.g., "三、VAV变风量空调机", "四、蓄冷(蓄热)设备")
        if (/^[一二三四五六七八九十]+、/.test(firstCol) && !firstCol.includes('·')) {
            subsections.push({row, text: firstCol});
        }
        
        // Find quota codes in entire row
        for (let col = 1; col <= 25; col++) {
            const value = getCellValue(row, col);
            if (/^\d+[A-Z]-\d+$/.test(value)) {
                quotaCodes.push({row, code: value});
                break; // Only take first code from each row
            }
        }
    }

    console.log('\n=== 找到的子节信息 ===');
    subsections.forEach(sub => {
        console.log(`行${sub.row}: ${sub.text}`);
    });

    console.log('\n=== 关键定额编号及其位置 ===');
    const targetCodes = ['1B-27', '1B-28', '1B-29', '1B-30', '1B-31', '1B-32', '1B-33', '1B-34', '1B-35', '1B-36'];
    const relevantCodes = quotaCodes.filter(q => targetCodes.includes(q.code));
    
    relevantCodes.forEach(q => {
        console.log(`行${q.row}: ${q.code}`);
    });

    console.log('\n=== 分析层级关系 ===');
    
    // Find the problematic subsections
    const vavSubsection = subsections.find(s => s.text.includes('VAV变风量空调机') || s.text.includes('三、'));
    const storageSubsection = subsections.find(s => s.text.includes('蓄冷') || s.text.includes('蓄热') || s.text.includes('四、'));
    
    if (vavSubsection && storageSubsection) {
        console.log(`VAV变风量空调机子节 在行${vavSubsection.row}: ${vavSubsection.text}`);
        console.log(`蓄冷(蓄热)设备子节 在行${storageSubsection.row}: ${storageSubsection.text}`);
        
        console.log('\n问题分析:');
        console.log(`编号 1B-28 到 1B-36 的位置:`);
        relevantCodes.forEach(q => {
            if (q.row < storageSubsection.row) {
                console.log(`  ${q.code} (行${q.row}) - 错误：在蓄冷(蓄热)设备子节(行${storageSubsection.row})之前`);
            } else {
                console.log(`  ${q.code} (行${q.row}) - 正确：在蓄冷(蓄热)设备子节(行${storageSubsection.row})之后`);
            }
        });
    }

    // Look for sub-sub sections and sub-sub-sub sections around relevant areas
    console.log('\n=== 查找缺失的子子节和子子子节 (第二章) ===');
    for (let row = 200; row <= 400; row++) {
        const firstCol = getCellValue(row, 1);
        
        // Look for numbered subsections (1., 2., etc.) and parenthetical subsections ((1), (2), etc.)
        if (/^\d+\.|^\(\d+\)|^\([一二三四五六七八九十]+\)/.test(firstCol)) {
            console.log(`行${row}: ${firstCol.substring(0, 60)}...`);
        }
    }
}

analyzeStructure().catch(console.error);