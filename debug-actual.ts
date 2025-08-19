import * as ExcelJS from 'exceljs';

async function findActualStructure() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('sample/input.xlsx');
    const worksheet = workbook.worksheets[0];

    console.log('=== 查找实际的章节和定额编号位置 ===');

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
                return String(cell.value.result || cell.value.formula);
            }
        }
        
        return String(cell.value).trim();
    }

    // Find chapter and section headers
    const headers: Array<{row: number, text: string, type: string}> = [];
    const quotaCodes: Array<{row: number, code: string}> = [];
    
    for (let row = 1; row <= worksheet.rowCount; row++) {
        const firstCol = getCellValue(row, 1);
        
        // Find chapter headers
        if (/^第[一二三四五六七八九十]+章/.test(firstCol)) {
            headers.push({row, text: firstCol, type: 'chapter'});
        }
        
        // Find section headers
        if (/^第[一二三四五六七八九十]+节/.test(firstCol)) {
            headers.push({row, text: firstCol, type: 'section'});
        }
        
        // Find subsection headers (like "一、减振装置安装", "二、柴油发电机组", etc.)
        // But skip table of contents (which have dots)
        if (/^[一二三四五六七八九十]+、[^·]/.test(firstCol)) {
            headers.push({row, text: firstCol.substring(0, 30), type: 'subsection'});
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

    console.log('\n=== 章节结构 ===');
    headers.forEach(h => {
        console.log(`${h.type.padEnd(12)} 行${h.row.toString().padStart(3)}: ${h.text}`);
    });

    console.log('\n=== 特定定额编号位置 ===');
    const targetCodes = ['1B-26', '1B-27', '1B-28', '1B-29', '1B-30', '1B-31', '1B-32', '1B-33', '1B-34', '1B-35', '1B-36'];
    const relevantCodes = quotaCodes.filter(q => targetCodes.includes(q.code));
    
    relevantCodes.forEach(q => {
        console.log(`行${q.row.toString().padStart(3)}: ${q.code}`);
    });

    // Analyze the hierarchical problem
    console.log('\n=== 层级问题分析 ===');
    
    const vavSubsection = headers.find(h => h.type === 'subsection' && (h.text.includes('VAV') || h.text.includes('三、')));
    const storageSubsection = headers.find(h => h.type === 'subsection' && (h.text.includes('蓄冷') || h.text.includes('蓄热') || h.text.includes('四、')));
    
    if (vavSubsection && storageSubsection) {
        console.log(`VAV变风量空调机子节: 行${vavSubsection.row} - ${vavSubsection.text}`);
        console.log(`蓄冷(蓄热)设备子节: 行${storageSubsection.row} - ${storageSubsection.text}`);
        
        console.log('\n编号归属分析:');
        relevantCodes.forEach(q => {
            let belongsTo = 'unknown';
            
            // Find which subsection this code should belong to
            const applicableHeaders = headers.filter(h => 
                h.type === 'subsection' && h.row < q.row
            ).sort((a, b) => b.row - a.row); // Get the closest preceding subsection
            
            if (applicableHeaders.length > 0) {
                belongsTo = applicableHeaders[0].text;
            }
            
            console.log(`  ${q.code} (行${q.row}) -> ${belongsTo}`);
        });
    }
}

findActualStructure().catch(console.error);