import { ExcelAnalyzer } from './excel-analyzer';

export class DebugAnalyzer {
  private analyzer: ExcelAnalyzer;

  constructor() {
    this.analyzer = new ExcelAnalyzer();
  }

  async analyzeFile(filePath: string): Promise<void> {
    await this.analyzer.loadFile(filePath);

    console.log('\n=== 调试分析开始 ===');

    // 分析前100行的内容结构
    this.analyzeRowStructure();

    // 寻找关键标识行
    this.findKeyRows();

    // 分析边框模式
    this.analyzeBorderPatterns();

    // 搜索附注信息模式
    this.searchNoteInfoPatterns();

    // 搜索子目信息模式
    this.searchSubItemPatterns();

    // 分析附注行周围的内容
    this.analyzeRowsAroundNotes();

    // 分析材料表结构
    this.analyzeMaterialTableStructure();
  }

  private analyzeRowStructure(): void {
    console.log('\n--- 行结构分析 ---');

    for (let row = 1; row <= Math.min(100, 839); row++) {
      const rowData: any[] = [];
      let hasContent = false;

      for (let col = 1; col <= Math.min(15, 32); col++) {
        const cellInfo = this.analyzer.getCellInfo(row, col);
        let value = cellInfo.value;

        // 处理复杂对象
        if (value && typeof value === 'object') {
          if ('richText' in value && Array.isArray(value.richText)) {
            value = value.richText.map((rt: any) => rt.text || '').join('');
          } else if ('text' in value) {
            value = value.text;
          } else {
            value = String(value);
          }
        }

        if (value && String(value).trim()) {
          hasContent = true;
        }

        rowData.push(String(value || '').substring(0, 15));
      }

      if (hasContent) {
        console.log(`行${row}: [${rowData.join(' | ')}]`);

        // 检查是否是潜在的标题行
        const rowText = rowData.join(' ').toLowerCase();
        if (rowText.includes('编号') || rowText.includes('名称') || rowText.includes('工作内容') ||
            rowText.includes('附注') || rowText.includes('子目') || rowText.includes('含量')) {
          console.log(`  -> 可能是标题行: ${rowText}`);
        }

        // 检查是否包含定额编号
        const hasDefineCode = rowData.some(cell => /\d+[A-Z]-\d+/.test(cell));
        if (hasDefineCode) {
          console.log(`  -> 包含定额编号`);
        }
      }

      // 只显示前50行和后段的关键行
      if (row === 50) {
        console.log('... 跳过中间行，显示后段重要行 ...');
        row = Math.max(50, 839 - 50);
      }
    }
  }

  private findKeyRows(): void {
    console.log('\n--- 关键行查找 ---');

    const keywords = ['工作内容', '附注信息', '子目信息', '含量表', '编号', '名称', '单位', '单价'];

    for (let row = 1; row <= 839; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= 32; col++) {
        const cellInfo = this.analyzer.getCellInfo(row, col);
        let value = cellInfo.value;

        if (value && typeof value === 'object') {
          if ('richText' in value && Array.isArray(value.richText)) {
            value = value.richText.map((rt: any) => rt.text || '').join('');
          } else if ('text' in value) {
            value = value.text;
          } else {
            value = String(value);
          }
        }

        rowData.push(String(value || ''));
      }

      const rowText = rowData.join(' ');

      for (const keyword of keywords) {
        if (rowText.includes(keyword)) {
          console.log(`行${row} 包含关键词"${keyword}": ${rowText.substring(0, 100)}...`);
          break;
        }
      }
    }
  }

  private analyzeBorderPatterns(): void {
    console.log('\n--- 边框模式分析 ---');

    let borderRegions = 0;
    let lastBorderRow = 0;

    for (let row = 1; row <= 839; row++) {
      let hasBorder = false;

      for (let col = 1; col <= 32; col++) {
        const cellInfo = this.analyzer.getCellInfo(row, col);
        if (cellInfo.borders.top || cellInfo.borders.bottom ||
            cellInfo.borders.left || cellInfo.borders.right) {
          hasBorder = true;
          break;
        }
      }

      if (hasBorder) {
        if (row - lastBorderRow > 50) {
          borderRegions++;
          console.log(`边框区域 ${borderRegions} 开始于行 ${row}`);
        }
        lastBorderRow = row;
      }
    }

    console.log(`总共发现 ${borderRegions} 个可能的边框区域`);
  }

    // 搜索附注信息模式
  searchNoteInfoPatterns(): void {
    console.log('\n--- 搜索附注信息模式 ---');

    for (let row = 1; row <= 839; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= 32; col++) {
        const cellInfo = this.analyzer.getCellInfo(row, col);
        const value = this.analyzer.processExcelValue(cellInfo.value);
        if (value && typeof value === 'string') {
          rowData.push(value);
        }
      }

      const rowText = rowData.join(' ');

      // 搜索包含"未包括"的行
      if (rowText.includes('未包括')) {
        console.log(`行${row}: ${rowText.substring(0, 100)}...`);
      }

      // 搜索包含"附注"的行
      if (rowText.includes('附注')) {
        console.log(`行${row}: ${rowText.substring(0, 100)}...`);
      }

      // 搜索包含"备注"的行
      if (rowText.includes('备注')) {
        console.log(`行${row}: ${rowText.substring(0, 100)}...`);
      }
    }
  }

  // 搜索子目信息模式
  searchSubItemPatterns(): void {
    console.log('\n--- 搜索子目信息模式 ---');

    for (let row = 1; row <= 839; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= Math.min(5, 32); col++) {
        const cellInfo = this.analyzer.getCellInfo(row, col);
        const value = this.analyzer.processExcelValue(cellInfo.value);
        if (value && typeof value === 'string') {
          rowData.push(value);
        }
      }

      const firstCol = rowData[0] || '';
      const secondCol = rowData[1] || '';

      // 搜索章节标题
      if (/^第[一二三四五六七八九十]+章/.test(firstCol)) {
        console.log(`行${row}: 章节标题 - ${firstCol} ${secondCol}`);
      }

      // 搜索节标题
      if (/^第[一二三四五六七八九十]+节/.test(firstCol)) {
        console.log(`行${row}: 节标题 - ${firstCol} ${secondCol}`);
      }

      // 搜索子节标题
      if (/^[一二三四五六七八九十]+、/.test(firstCol)) {
        console.log(`行${row}: 子节标题 - ${firstCol} ${secondCol}`);
      }

      // 搜索数字子节标题
      if (/^\d+、/.test(firstCol)) {
        console.log(`行${row}: 数字子节标题 - ${firstCol} ${secondCol}`);
      }

      // 搜索括号子节标题
      if (/^\([一二三四五六七八九十]+\)/.test(firstCol)) {
        console.log(`行${row}: 括号子节标题 - ${firstCol} ${secondCol}`);
      }

      // 搜索定额编号
      if (/^\d+[A-Z]-\d+$/.test(firstCol)) {
        console.log(`行${row}: 定额编号 - ${firstCol} ${secondCol}`);
      }
    }
  }

  // 分析特定行周围的内容
  analyzeRowsAroundNotes(): void {
    console.log('\n--- 分析附注行周围的内容 ---');

    const noteRows = [301, 310, 323, 342, 352, 361, 372, 382, 392, 412, 479];

    for (const noteRow of noteRows) {
      console.log(`\n=== 分析行 ${noteRow} 周围的内容 ===`);

      // 显示前5行和后5行
      for (let row = Math.max(1, noteRow - 5); row <= Math.min(this.analyzer.worksheet.rowCount, noteRow + 5); row++) {
        const rowData: string[] = [];

        for (let col = 1; col <= Math.min(10, this.analyzer.worksheet.columnCount); col++) {
          const cellInfo = this.analyzer.getCellInfo(row, col);
          const value = this.analyzer.processExcelValue(cellInfo.value);
          rowData.push(String(value || ''));
        }

        const rowText = rowData.join(' | ');
        const marker = row === noteRow ? '>>> ' : '    ';
        console.log(`${marker}行${row}: ${rowText.substring(0, 150)}...`);
      }
    }
  }

  // 分析材料表结构
  analyzeMaterialTableStructure(): void {
    console.log('\n--- 分析材料表结构 ---');

    const materialTableRows = [108, 288, 297, 307, 318, 339, 358, 441, 476];

    for (const tableRow of materialTableRows) {
      console.log(`\n=== 分析材料表行 ${tableRow} ===`);

      // 显示表头行
      console.log('表头行:');
      for (let col = 1; col <= Math.min(15, this.analyzer.worksheet.columnCount); col++) {
        const cellInfo = this.analyzer.getCellInfo(tableRow, col);
        const value = this.analyzer.processExcelValue(cellInfo.value);
        console.log(`  列${col}: "${String(value || '')}"`);
      }

      // 显示数据行
      console.log('数据行:');
      for (let dataRow = tableRow + 1; dataRow <= Math.min(tableRow + 10, this.analyzer.worksheet.rowCount); dataRow++) {
        const rowData: string[] = [];
        for (let col = 1; col <= Math.min(15, this.analyzer.worksheet.columnCount); col++) {
          const cellInfo = this.analyzer.getCellInfo(dataRow, col);
          const value = this.analyzer.processExcelValue(cellInfo.value);
          rowData.push(String(value || ''));
        }

        const rowText = rowData.join(' | ');
        if (rowText.trim()) {
          console.log(`  行${dataRow}: ${rowText.substring(0, 200)}...`);
        }
      }
    }
  }
}

// 如果直接运行此文件
if (require.main === module) {
  const analyzer = new DebugAnalyzer();
  analyzer.analyzeFile('./sample/input.xlsx').catch(console.error);
}
