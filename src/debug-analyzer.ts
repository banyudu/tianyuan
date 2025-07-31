import * as ExcelJS from 'exceljs';
import * as path from 'path';

async function analyzeInputFile() {
  console.log('=== 分析输入文件结构 ===');
  
  const workbook = new ExcelJS.Workbook();
  const inputFile = path.join(__dirname, '../sample/input.xlsx');
  
  await workbook.xlsx.readFile(inputFile);
  const worksheet = workbook.worksheets[0];
  
  console.log(`工作表名称: ${worksheet.name}`);
  console.log(`行数: ${worksheet.rowCount}, 列数: ${worksheet.columnCount}`);
  
  // 分析前50行的结构
  console.log('\n=== 前50行数据结构分析 ===');
  for (let row = 1; row <= Math.min(50, worksheet.rowCount); row++) {
    const rowData: string[] = [];
    let hasContent = false;
    
    for (let col = 1; col <= Math.min(15, worksheet.columnCount); col++) {
      const cell = worksheet.getCell(row, col);
      let value = '';
      
      if (cell.value !== null && cell.value !== undefined) {
        if (typeof cell.value === 'object') {
          if ('richText' in cell.value && Array.isArray(cell.value.richText)) {
            value = cell.value.richText.map((rt: any) => rt.text || '').join('');
          } else if ('text' in cell.value) {
            value = cell.value.text;
          } else if ('result' in cell.value) {
            value = cell.value.result;
          } else if ('formula' in cell.value) {
            value = cell.value.result || cell.value.formula;
          } else {
            value = String(cell.value);
          }
        } else {
          value = String(cell.value);
        }
        
        if (value.trim()) {
          hasContent = true;
        }
      }
      
      rowData.push(value);
    }
    
    if (hasContent) {
      const rowText = rowData.join(' | ').trim();
      if (rowText) {
        console.log(`行${row}: ${rowText.substring(0, 150)}${rowText.length > 150 ? '...' : ''}`);
      }
    }
  }
  
  // 查找定额编号模式
  console.log('\n=== 定额编号模式分析 ===');
  const quotaCodes: {row: number, col: number, code: string}[] = [];
  
  for (let row = 1; row <= worksheet.rowCount && quotaCodes.length < 20; row++) {
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = worksheet.getCell(row, col);
      if (cell.value) {
        const str = String(cell.value).trim();
        if (/^\d+[A-Z]-\d+$/.test(str)) {
          quotaCodes.push({row, col, code: str});
        }
      }
    }
  }
  
  console.log(`找到 ${quotaCodes.length} 个定额编号示例:`);
  quotaCodes.slice(0, 10).forEach(item => {
    console.log(`  行${item.row}, 列${item.col}: ${item.code}`);
  });
  
  // 查找章节标题
  console.log('\n=== 章节标题模式分析 ===');
  const chapters: {row: number, text: string}[] = [];
  
  for (let row = 1; row <= worksheet.rowCount && chapters.length < 10; row++) {
    const cell = worksheet.getCell(row, 1);
    if (cell.value) {
      const str = String(cell.value).trim();
      if (/第[一二三四五六七八九十]+章/.test(str)) {
        chapters.push({row, text: str});
      }
    }
  }
  
  console.log(`找到 ${chapters.length} 个章节标题:`);
  chapters.forEach(item => {
    console.log(`  行${item.row}: ${item.text}`);
  });
  
  // 查找工作内容模式
  console.log('\n=== 工作内容模式分析 ===');
  const workContents: {row: number, text: string}[] = [];
  
  for (let row = 1; row <= worksheet.rowCount && workContents.length < 10; row++) {
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = worksheet.getCell(row, col);
      if (cell.value) {
        const str = String(cell.value).trim();
        if (str.includes('工作内容') && str.includes('：')) {
          const preview = str.substring(0, 100);
          workContents.push({row, text: preview});
          break;
        }
      }
    }
  }
  
  console.log(`找到 ${workContents.length} 个工作内容:`);
  workContents.forEach(item => {
    console.log(`  行${item.row}: ${item.text}...`);
  });
  
  // 查找附注信息模式
  console.log('\n=== 附注信息模式分析 ===');
  const notes: {row: number, text: string}[] = [];
  
  for (let row = 1; row <= worksheet.rowCount && notes.length < 10; row++) {
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = worksheet.getCell(row, col);
      if (cell.value) {
        const str = String(cell.value).trim();
        if (str.includes('注 : 未包括') || str.includes('注: 未包括')) {
          const preview = str.substring(0, 100);
          notes.push({row, text: preview});
          break;
        }
      }
    }
  }
  
  console.log(`找到 ${notes.length} 个附注信息:`);
  notes.forEach(item => {
    console.log(`  行${item.row}: ${item.text}...`);
  });
  
  // 查找材料表头
  console.log('\n=== 材料表头模式分析 ===');
  const materialHeaders: {row: number, text: string}[] = [];
  
  for (let row = 1; row <= worksheet.rowCount && materialHeaders.length < 10; row++) {
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = worksheet.getCell(row, col);
      if (cell.value) {
        const str = String(cell.value).trim();
        if (str.includes('人材机名称') || (str.includes('人') && str.includes('材') && str.includes('机') && str.includes('名称'))) {
          const preview = str.substring(0, 150);
          materialHeaders.push({row, text: preview});
          break;
        }
      }
    }
  }
  
  console.log(`找到 ${materialHeaders.length} 个材料表头:`);
  materialHeaders.forEach(item => {
    console.log(`  行${item.row}: ${item.text}...`);
  });
}

// 运行分析
analyzeInputFile().catch(console.error);