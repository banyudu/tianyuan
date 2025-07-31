import * as ExcelJS from 'exceljs';

export interface CellInfo {
  address: string;
  value: any;
  merged: boolean;
  borders: {
    top?: boolean;
    bottom?: boolean;
    left?: boolean;
    right?: boolean;
  };
  row: number;
  col: number;
}

export interface TableRegion {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  cells: CellInfo[];
  metadata?: {
    centerRow: number;
    codes: string[];
    workContentRow: number;
    materialDataRows: number[];
    type: string;
  };
}

export class ExcelAnalyzer {
  private readonly workbook: ExcelJS.Workbook;
  worksheet!: ExcelJS.Worksheet;

  constructor() {
    this.workbook = new ExcelJS.Workbook();
  }

  async loadFile(filePath: string): Promise<void> {
    await this.workbook.xlsx.readFile(filePath);
    // 假设数据在第一个工作表中
    this.worksheet = this.workbook.worksheets[0];
    console.log(`加载文件成功: ${filePath}`);
    console.log(`工作表名称: ${this.worksheet.name}`);
    console.log(`行数: ${this.worksheet.rowCount}, 列数: ${this.worksheet.columnCount}`);
  }

  // 获取单元格信息，包括边框样式
  getCellInfo(row: number, col: number): CellInfo {
    const cell = this.worksheet.getCell(row, col);

    // 检查边框样式
    const borders = {
      top: Boolean(cell.border?.top?.style ?? false),
      bottom: Boolean(cell.border?.bottom?.style ?? false),
      left: Boolean(cell.border?.left?.style ?? false),
      right: Boolean(cell.border?.right?.style ?? false),
    };

    return {
      address: cell.address,
      value: cell.value,
      merged: cell.isMerged,
      borders,
      row,
      col,
    };
  }

  // 扫描整个工作表，识别有边框的区域
  scanBorderedRegions(): TableRegion[] {
    const regions: TableRegion[] = [];
    const processedCells = new Set<string>();

    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellKey = `${row},${col}`;
        if (processedCells.has(cellKey)) continue;

        const cellInfo = this.getCellInfo(row, col);

        // 如果发现有边框的单元格，尝试识别完整的表格区域
        if (this.hasBorder(cellInfo)) {
          const region = this.identifyTableRegion(row, col, processedCells);
          if (region && region.cells.length > 0) {
            regions.push(region);
          }
        }
      }
    }

    // 过滤和验证区域
    return this.filterValidDataRegions(regions);
  }

  // 过滤出有效的数据区域
  private filterValidDataRegions(regions: TableRegion[]): TableRegion[] {
    const validRegions: TableRegion[] = [];

    for (const region of regions) {
      // 检查区域是否包含有效数据
      if (this.isValidDataRegion(region)) {
        validRegions.push(region);
        console.log(`有效数据区域: 行${region.startRow}-${region.endRow}, 类型: ${this.identifyDataType(region)}`);
      } else {
        console.log(`跳过无效区域: 行${region.startRow}-${region.endRow} (可能是目录页)`);
      }
    }

    return validRegions;
  }

  // 检查区域是否包含有效数据
  private isValidDataRegion(region: TableRegion): boolean {
    const data = this.getRegionData(region);

    // 检查是否包含定额编号
    const hasDefineCode = data.some(row =>
      row.some(cell => {
        const str = String(cell || '').trim();
        return /^\d+[A-Z]-\d+$/.test(str);
      })
    );

    // 检查是否包含关键标题
    const hasKeyTitles = data.some(row => {
      const rowText = row.join(' ').toLowerCase();
      return rowText.includes('工作内容') ||
             rowText.includes('人材机名称') ||
             rowText.includes('单位') ||
             (rowText.includes('编号') && rowText.includes('名称'));
    });

    // 排除目录页特征
    const hasTableOfContents = data.some(row => {
      const rowText = row.join(' ');
      return rowText.includes('····') || // 目录页的点线
             rowText.includes('第一章') ||
             rowText.includes('第二章') ||
             (rowText.includes('页') && rowText.match(/\d+$/)); // 以页码结尾
    });

    // 必须有定额编号或关键标题，且不能是目录页
    return (hasDefineCode || hasKeyTitles) && !hasTableOfContents;
  }

  // 检查单元格是否有边框
  private hasBorder(cellInfo: CellInfo): boolean {
    return (
      Boolean(cellInfo.borders.top ?? false) ||
      Boolean(cellInfo.borders.bottom ?? false) ||
      Boolean(cellInfo.borders.left ?? false) ||
      Boolean(cellInfo.borders.right ?? false)
    );
  }

  // 识别从指定单元格开始的表格区域
  private identifyTableRegion(
    startRow: number,
    startCol: number,
    processedCells: Set<string>
  ): TableRegion | null {
    const regionCells: CellInfo[] = [];
    let minRow = startRow,
      maxRow = startRow;
    let minCol = startCol,
      maxCol = startCol;

    // 使用广度优先搜索找到连续的有边框区域
    const queue: [number, number][] = [[startRow, startCol]];
    const visited = new Set<string>();

    while (queue.length > 0) {
      const [row, col] = queue.shift()!;
      const cellKey = `${row},${col}`;

      if (
        visited.has(cellKey) ||
        row < 1 ||
        row > this.worksheet.rowCount ||
        col < 1 ||
        col > this.worksheet.columnCount
      ) {
        continue;
      }

      const cellInfo = this.getCellInfo(row, col);

      // 如果单元格有边框或有内容，加入区域
      if (this.hasBorder(cellInfo) || (cellInfo.value !== null && cellInfo.value !== undefined)) {
        visited.add(cellKey);
        processedCells.add(cellKey);
        regionCells.push(cellInfo);

        minRow = Math.min(minRow, row);
        maxRow = Math.max(maxRow, row);
        minCol = Math.min(minCol, col);
        maxCol = Math.max(maxCol, col);

        // 检查相邻单元格
        queue.push([row - 1, col], [row + 1, col], [row, col - 1], [row, col + 1]);
      }
    }

    if (regionCells.length === 0) return null;

    return {
      startRow: minRow,
      endRow: maxRow,
      startCol: minCol,
      endCol: maxCol,
      cells: regionCells,
    };
  }

  // 获取区域内的数据为二维数组
  getRegionData(region: TableRegion): any[][] {
    const data: any[][] = [];

    for (let row = region.startRow; row <= region.endRow; row++) {
      const rowData: any[] = [];
      for (let col = region.startCol; col <= region.endCol; col++) {
        const cell = this.worksheet.getCell(row, col);
        let value = cell.value;

        // 处理合并单元格 - 获取主单元格的值
        if (cell.isMerged) {
          const masterCell = this.worksheet.getCell(cell.master?.address || cell.address);
          value = masterCell.value;
        }

        // 更好地处理特殊值类型
        value = this.processExcelValue(value);

        rowData.push(value || '');
      }
      data.push(rowData);
    }

    return data;
  }

  // 处理Excel中的特殊值类型
  processExcelValue(value: any): string | number | boolean | null {
    if (value === null || value === undefined) {
      return null;
    }

    // 处理富文本对象
    if (value && typeof value === 'object') {
      if ('richText' in value && Array.isArray(value.richText)) {
        return value.richText.map((rt: any) => rt.text || '').join('');
      }
      if ('text' in value) {
        return value.text;
      }
      if ('result' in value) {
        return value.result;
      }
      if ('formula' in value) {
        return value.result || value.formula;
      }
      // 处理超链接
      if ('hyperlink' in value && 'text' in value) {
        return value.text;
      }
      // 如果是其他对象类型，尝试转换为字符串
      if (typeof value.toString === 'function') {
        const str = value.toString();
        return str === '[object Object]' ? '' : str;
      }
      return '';
    }

    return value;
  }

  // 打印区域信息用于调试
  printRegionInfo(regions: TableRegion[]): void {
    console.log(`\n发现 ${regions.length} 个有效表格区域:`);
    regions.forEach((region, index) => {
      console.log(`\n区域 ${index + 1}:`);
      console.log(
        `  位置: 行${region.startRow}-${region.endRow}, 列${region.startCol}-${region.endCol}`
      );
      console.log(
        `  大小: ${region.endRow - region.startRow + 1} x ${region.endCol - region.startCol + 1}`
      );
      console.log(`  单元格数量: ${region.cells.length}`);

      // 显示前几行数据作为预览
      const data = this.getRegionData(region);
      console.log(`  数据预览:`);
      data.slice(0, 3).forEach((row, rowIndex) => {
        console.log(
          `    行${rowIndex + 1}: [${row
            .slice(0, 5)
            .map(cell => String(cell).substring(0, 10))
            .join(', ')}]`
        );
      });

      // 识别数据类型
      const dataType = this.identifyDataType(region);
      console.log(`  识别类型: ${dataType}`);
    });
  }

  // 根据内容特征识别数据类型
  identifyDataType(region: TableRegion): string {
    const data = this.getRegionData(region);

    // 检查前几行是否包含关键字
    let allText = '';
    for (let i = 0; i < Math.min(10, data.length); i++) {
      allText += data[i].join(' ').toLowerCase() + ' ';
    }

    // 更精确的关键词匹配
    if (allText.includes('工作内容') && allText.includes('一、工作内容')) {
      return '工作内容';
    }
    if (allText.includes('附注') || allText.includes('备注') || allText.includes('未包括')) {
      return '附注信息';
    }
    if (allText.includes('人材机名称') || (allText.includes('人') && allText.includes('材') && allText.includes('机'))) {
      return '含量表';
    }
    if (allText.includes('子目') && (allText.includes('名称') || allText.includes('定额'))) {
      return '子目信息';
    }

    // 根据数据模式判断
    if (data.length > 5) {
      // 检查是否有定额编号模式
      const hasDefineCode = data.some(row =>
        row.some(cell => {
          const str = String(cell || '');
          return /^\d+[A-Z]-\d+$/.test(str);
        })
      );

      if (hasDefineCode) {
        // 检查列数来区分含量表和其他
        if (data[0]?.length >= 8) {
          return '含量表';
        } else if (data[0]?.length <= 3) {
          // 检查是否有工作关键词
          const hasWorkKeywords = data.some(row => {
            const text = row.join(' ').toLowerCase();
            return text.includes('安装') || text.includes('测位') || text.includes('切管') || text.includes('划线');
          });

          return hasWorkKeywords ? '工作内容' : '附注信息';
        } else {
          return '子目信息';
        }
      }
    }

    return '未知类型';
  }

  // 查找特定数据类型的区域
  findRegionsByType(regions: TableRegion[], targetType: string): TableRegion[] {
    return regions.filter(region => this.identifyDataType(region) === targetType);
  }

  // 直接搜索数据区域（用于处理边框识别失败的情况）
  scanDirectDataRegions(): TableRegion[] {
    const regions: TableRegion[] = [];

    // 寻找"一、工作内容："标识
    for (let row = 700; row <= this.worksheet.rowCount; row++) {
      const rowData: string[] = [];
      for (let col = 1; col <= 5; col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        rowData.push(String(value || ''));
      }

      const rowText = rowData.join(' ');
      if (rowText.includes('一、工作内容')) {
        console.log(`找到工作内容标识行: ${row}`);

        // 创建一个从这一行开始到文件末尾的区域
        const region: TableRegion = {
          startRow: row,
          endRow: this.worksheet.rowCount,
          startCol: 1,
          endCol: Math.min(15, this.worksheet.columnCount),
          cells: []
        };

        regions.push(region);
        break;
      }
    }

    return regions;
  }

  // 基于定额编号识别数据区域
  scanDefineCodeRegions(): TableRegion[] {
    const regions: TableRegion[] = [];

    console.log('\n=== 基于定额编号识别数据区域 ===');

    // 首先找到所有包含定额编号的行
    const codeRows: {row: number, codes: string[]}[] = [];

    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      const codes: string[] = [];

      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        const str = String(value || '').trim();

        if (/^\d+[A-Z]-\d+$/.test(str)) {
          codes.push(str);
        }
      }

      if (codes.length > 0) {
        const uniqueCodes = [...new Set(codes)];
        codeRows.push({row, codes: uniqueCodes});
      }
    }

    console.log(`找到 ${codeRows.length} 行包含定额编号`);

    // 为每个定额编号行创建一个数据区域
    for (const codeRow of codeRows) {
      // 寻找该行附近的相关数据
      const region = this.createRegionAroundRow(codeRow.row, codeRow.codes);
      if (region) {
        regions.push(region);
      }
    }

    return regions;
  }

  // 在指定行周围创建数据区域
  private createRegionAroundRow(centerRow: number, codes: string[]): TableRegion | null {
    // 向前向后扩展几行，寻找相关的工作内容和材料数据
    const startRow = Math.max(1, centerRow - 5);
    const endRow = Math.min(this.worksheet.rowCount, centerRow + 15);

    // 寻找相关的工作内容行
    let workContentRow = -1;
    let materialDataRows: number[] = [];

    for (let row = startRow; row <= endRow; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= Math.min(10, this.worksheet.columnCount); col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        rowData.push(String(value || ''));
      }

      const rowText = rowData.join(' ').toLowerCase();

      // 检查是否是工作内容行
      if (rowText.includes('工作') && rowText.includes('内容') && rowText.includes('：')) {
        workContentRow = row;
      }

      // 检查是否是材料数据行（包含"人材机名称"等）
      if (rowText.includes('人') && rowText.includes('材') && rowText.includes('机') && rowText.includes('名称')) {
        materialDataRows.push(row);
      }
    }

    // 创建一个包含相关数据的区域
    const region: TableRegion = {
      startRow: startRow,
      endRow: endRow,
      startCol: 1,
      endCol: Math.min(15, this.worksheet.columnCount),
      cells: [],
      metadata: {
        centerRow,
        codes,
        workContentRow,
        materialDataRows,
        type: this.determineRegionType(centerRow, workContentRow, materialDataRows)
      }
    };

    return region;
  }

  // 确定区域类型
  private determineRegionType(centerRow: number, workContentRow: number, materialDataRows: number[]): string {
    if (workContentRow > 0) {
      return '工作内容';
    } else if (materialDataRows.length > 0) {
      return '含量表';
    } else {
      return '子目信息';
    }
  }

  // 获取指定行的附注信息
  getNoteInfoFromRow(row: number, codes: string[]): {编号: string, 附注信息: string} | null {
    const rowData: string[] = [];

    for (let col = 1; col <= this.worksheet.columnCount; col++) {
      const cellInfo = this.getCellInfo(row, col);
      const value = this.processExcelValue(cellInfo.value);
      rowData.push(String(value || ''));
    }

    const rowText = rowData.join(' ');

    // 查找附注信息 - 通常在"附注："或"备注："后面
    const noteMatch = rowText.match(/(?:附注|备注)[：:](.*?)(?:\s*$)/);
    if (noteMatch) {
      const noteText = noteMatch[1].trim();
      if (noteText && codes.length > 0) {
        return {
          编号: codes[0],
          附注信息: noteText
        };
      }
    }

    // 查找"注 : 未包括"模式
    const noteMatch2 = rowText.match(/注\s*:\s*(.*?)(?:\s*$)/);
    if (noteMatch2) {
      const noteText = noteMatch2[1].trim();
      if (noteText && codes.length > 0) {
        return {
          编号: codes[0],
          附注信息: noteText
        };
      }
    }

    return null;
  }

  // 扫描整个工作表获取所有附注信息
  scanAllNoteInfo(): {编号: string, 附注信息: string}[] {
    const noteInfoList: {编号: string, 附注信息: string}[] = [];
    let debugCount = 0;

    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        rowData.push(String(value || ''));
      }

      const rowText = rowData.join(' ');

      // 查找包含"注 : 未包括"的行
      if (rowText.includes('注 : 未包括') || rowText.includes('注: 未包括')) {
        debugCount++;
        console.log(`DEBUG: 找到附注行 ${row}: ${rowText.substring(0, 100)}...`);

        // 向下查找最近的定额编号（通常在接下来的几行中）
        let nearestCode = '';
        for (let searchRow = row + 1; searchRow <= Math.min(this.worksheet.rowCount, row + 10); searchRow++) {
          for (let col = 1; col <= this.worksheet.columnCount; col++) {
            const cellInfo = this.getCellInfo(searchRow, col);
            const value = this.processExcelValue(cellInfo.value);
            if (value && typeof value === 'string') {
              const str = String(value).trim();
              if (/^\d+[A-Z]-\d+$/.test(str)) {
                nearestCode = str;
                console.log(`DEBUG: 找到关联定额编号: ${nearestCode} (行${searchRow}, 列${col})`);
                break;
              }
            }
          }
          if (nearestCode) break;
        }

        if (nearestCode) {
          // 提取附注信息
          const noteMatch = rowText.match(/注\s*:\s*(.*?)(?:\s*$)/);
          if (noteMatch) {
            const noteText = noteMatch[1].trim();
            // 清理重复的文本，只保留第一个"注 : "后面的内容
            const cleanNoteText = noteText.replace(/\s*注\s*:\s*.*$/g, '').trim();
            noteInfoList.push({
              编号: nearestCode,
              附注信息: cleanNoteText
            });
            console.log(`DEBUG: 添加附注信息: ${nearestCode} -> ${cleanNoteText.substring(0, 50)}...`);
          }
        } else {
          console.log(`DEBUG: 未找到关联定额编号，跳过行 ${row}`);
        }
      }
    }

    console.log(`DEBUG: 总共扫描了 ${this.worksheet.rowCount} 行，找到 ${debugCount} 行包含"注 : 未包括"`);
    return noteInfoList;
  }

  // 扫描整个工作表获取所有子目信息
  scanAllSubItemInfo(): any[] {
    const subItemList: any[] = [];
    let currentChapter = '';
    let currentSection = '';
    let currentSubsection = '';
    let currentSubSubsection = '';
    let currentItem = '';
    let debugCount = 0;

    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= Math.min(15, this.worksheet.columnCount); col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        rowData.push(String(value || ''));
      }

      const firstCol = rowData[0] || '';
      const secondCol = rowData[1] || '';

      // 检查章节标题
      if (/^第[一二三四五六七八九十]+章/.test(firstCol)) {
        currentChapter = firstCol.replace(/^第([一二三四五六七八九十]+)章.*/, '第$1章');
        currentSection = '';
        currentSubsection = '';
        currentSubSubsection = '';
        currentItem = '';
        console.log(`DEBUG: 章节标题 ${row}: ${currentChapter}`);
      }

      // 检查节标题
      if (/^第[一二三四五六七八九十]+节/.test(firstCol)) {
        currentSection = firstCol.replace(/^第([一二三四五六七八九十]+)节.*/, '第$1节');
        currentSubsection = '';
        currentSubSubsection = '';
        currentItem = '';
        console.log(`DEBUG: 节标题 ${row}: ${currentSection}`);
      }

      // 检查子节标题
      if (/^[一二三四五六七八九十]+、/.test(firstCol)) {
        currentSubsection = firstCol.replace(/^([一二三四五六七八九十]+)、.*/, '$1、');
        currentSubSubsection = '';
        currentItem = '';
        console.log(`DEBUG: 子节标题 ${row}: ${currentSubsection}`);
      }

      // 检查数字子节标题
      if (/^\d+、/.test(firstCol)) {
        currentSubSubsection = firstCol.replace(/^(\d+)、.*/, '$1、');
        currentItem = '';
        console.log(`DEBUG: 数字子节标题 ${row}: ${currentSubSubsection}`);
      }

      // 检查括号子节标题
      if (/^\([一二三四五六七八九十]+\)/.test(firstCol)) {
        currentItem = firstCol.replace(/^\(([一二三四五六七八九十]+)\).*/, '($1)');
        console.log(`DEBUG: 括号子节标题 ${row}: ${currentItem}`);
      }

      // 检查定额编号 - 扫描整行寻找定额编号
      let quotaCode = '';
      let quotaName = '';
      let priceData = {
        基价: 0,
        人工: 0,
        材料: 0,
        机械: 0,
        管理费: 0,
        利润: 0,
        其他: 0
      };

      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        if (value && typeof value === 'string') {
          const str = String(value).trim();
          if (/^\d+[A-Z]-\d+$/.test(str)) {
            quotaCode = str;

            // 更智能的名称提取：在定额编号周围搜索名称
            quotaName = this.findQuotaName(row, col);

            // 尝试从后续列提取价格数据
            priceData = this.extractPriceData(row, col);
            break;
          }
        }
      }

      if (quotaCode) {
        debugCount++;
        const symbol = this.determineSymbol(currentChapter, currentSection, currentSubsection, currentSubSubsection, currentItem);
        console.log(`DEBUG: 定额编号 ${row}: ${quotaCode} - ${quotaName} (符号: ${symbol})`);

        subItemList.push({
          符号: symbol,
          定额号: quotaCode,
          子目名称: quotaName || quotaCode, // 如果没有找到名称，使用定额号作为名称
          基价: priceData.基价,
          人工: priceData.人工,
          材料: priceData.材料,
          机械: priceData.机械,
          管理费: priceData.管理费,
          利润: priceData.利润,
          其他: priceData.其他,
          图片名称: ''
        });
      }
    }

    console.log(`DEBUG: 总共找到 ${debugCount} 个定额编号`);
    return subItemList;
  }

  // 提取价格数据的辅助方法
  private extractPriceData(row: number, quotaCol: number): any {
    const priceData = {
      基价: 0,
      人工: 0,
      材料: 0,
      机械: 0,
      管理费: 0,
      利润: 0,
      其他: 0
    };

    // 从定额编号列开始，向后扫描价格数据
    for (let col = quotaCol + 2; col <= Math.min(quotaCol + 15, this.worksheet.columnCount); col++) {
      const cellInfo = this.getCellInfo(row, col);
      const value = this.processExcelValue(cellInfo.value);

      if (value !== null && value !== undefined && value !== '') {
        const numValue = this.parseNumber(value);

        // 只处理有效的数字值（大于0且不是测试数据）
        if (numValue > 0 && numValue !== 1 && numValue !== 2 && numValue !== 3 &&
            numValue !== 5 && numValue !== 6 && numValue !== 7) {

          // 根据列位置分配价格类型（这是一个简化的逻辑）
          const colIndex = col - quotaCol;
          switch (colIndex) {
            case 2:
              priceData.基价 = numValue;
              break;
            case 3:
              priceData.人工 = numValue;
              break;
            case 4:
              priceData.材料 = numValue;
              break;
            case 5:
              priceData.机械 = numValue;
              break;
            case 6:
              priceData.管理费 = numValue;
              break;
            case 7:
              priceData.利润 = numValue;
              break;
            case 8:
              priceData.其他 = numValue;
              break;
          }
        }
      }
    }

    // 如果没有找到有效的价格数据，尝试从相邻行查找
    if (priceData.基价 === 0) {
      // 检查下一行是否有价格数据
      for (let searchRow = row + 1; searchRow <= Math.min(row + 3, this.worksheet.rowCount); searchRow++) {
        for (let col = quotaCol + 2; col <= Math.min(quotaCol + 15, this.worksheet.columnCount); col++) {
          const cellInfo = this.getCellInfo(searchRow, col);
          const value = this.processExcelValue(cellInfo.value);

          if (value !== null && value !== undefined && value !== '') {
            const numValue = this.parseNumber(value);

            if (numValue > 0 && numValue !== 1 && numValue !== 2 && numValue !== 3 &&
                numValue !== 5 && numValue !== 6 && numValue !== 7) {

              const colIndex = col - quotaCol;
              switch (colIndex) {
                case 2:
                  priceData.基价 = numValue;
                  break;
                case 3:
                  priceData.人工 = numValue;
                  break;
                case 4:
                  priceData.材料 = numValue;
                  break;
                case 5:
                  priceData.机械 = numValue;
                  break;
                case 6:
                  priceData.管理费 = numValue;
                  break;
                case 7:
                  priceData.利润 = numValue;
                  break;
                case 8:
                  priceData.其他 = numValue;
                  break;
              }
            }
          }
        }

        // 如果找到了有效数据，停止搜索
        if (priceData.基价 > 0 || priceData.人工 > 0 || priceData.材料 > 0) {
          break;
        }
      }
    }

    return priceData;
  }

  // 确定符号层级
  private determineSymbol(chapter: string, section: string, subsection: string, subSubsection: string, item: string): string {
    if (chapter && !section && !subsection && !subSubsection && !item) {
      return '$';
    } else if (chapter && section && !subsection && !subSubsection && !item) {
      return '$$';
    } else if (chapter && section && subsection && !subSubsection && !item) {
      return '$$$';
    } else if (chapter && section && subsection && subSubsection && !item) {
      return '$$$$';
    } else if (chapter && section && subsection && subSubsection && item) {
      return '$$$$$';
    } else {
      return '';
    }
  }

  // 获取指定行的工作内容
  getWorkContentFromRow(row: number, codes: string[]): {编号: string, 工作内容: string} | null {
    const rowData: string[] = [];

    for (let col = 1; col <= this.worksheet.columnCount; col++) {
      const cellInfo = this.getCellInfo(row, col);
      const value = this.processExcelValue(cellInfo.value);
      rowData.push(String(value || ''));
    }

    const rowText = rowData.join(' ');

    // 查找工作内容
    const workMatch = rowText.match(/工作\s*内容[：:](.*?)(?:\s*单位|$)/);
    if (workMatch) {
      let 工作内容 = workMatch[1].trim();

      // 数据清洗：去除重复的"工作内容"文本
      工作内容 = 工作内容.replace(/工作\s*内容[：:]\s*/g, '');

      // 去除重复的文本片段
      工作内容 = this.removeDuplicateText(工作内容);

      // 去除多余的空格和换行
      工作内容 = 工作内容.replace(/\s+/g, ' ').trim();

      const 编号 = codes.join(',');

      if (工作内容 && 编号 && 工作内容.length > 5) {
        return { 编号, 工作内容 };
      }
    }

    return null;
  }

  // 去除重复文本的辅助方法
  private removeDuplicateText(text: string): string {
    if (!text || text.length < 10) return text;

    // 去除重复的"工作内容"文本
    text = text.replace(/工作\s*内容[：:]\s*/g, '');

    // 去除重复的"一、"文本
    text = text.replace(/(一、)+/g, '一、');

    // 分割文本为句子
    const sentences = text.split(/[。，；]/).filter(s => s.trim().length > 0);

    // 去除重复的句子
    const uniqueSentences = [];
    const seen = new Set<string>();

    for (const sentence of sentences) {
      const cleanSentence = sentence.trim();
      if (cleanSentence && !seen.has(cleanSentence) && cleanSentence.length > 2) {
        // 进一步清理重复的短语
        const words = cleanSentence.split(/[、，\s]+/);
        const uniqueWords = [];
        const seenWords = new Set<string>();

        for (const word of words) {
          const cleanWord = word.trim();
          if (cleanWord && !seenWords.has(cleanWord) && cleanWord.length > 1) {
            uniqueWords.push(cleanWord);
            seenWords.add(cleanWord);
          }
        }

        const deduplicatedSentence = uniqueWords.join('、');
        if (deduplicatedSentence && !seen.has(deduplicatedSentence)) {
          uniqueSentences.push(deduplicatedSentence);
          seen.add(deduplicatedSentence);
        }
      }
    }

    return uniqueSentences.join('，') + '。';
  }

  // 获取指定行周围的材料数据
  getMaterialDataFromRegion(region: TableRegion): any[] {
    const materials: any[] = [];
    const data = this.getRegionData(region);

    // 寻找材料数据的开始行
    let dataStartIndex = -1;
    for (let i = 0; i < data.length; i++) {
      const rowText = data[i].join(' ').toLowerCase();
      if (rowText.includes('人') && rowText.includes('材') && rowText.includes('机')) {
        dataStartIndex = i + 1; // 下一行开始是数据
        break;
      }
    }

    if (dataStartIndex >= 0) {
      for (let i = dataStartIndex; i < data.length; i++) {
        const row = data[i];

        // 检查是否是有效的材料数据行
        if (row.length >= 6 && row[0] && row[1]) {
          const 编号 = String(row[0]).trim();
          const 名称 = String(row[1]).trim();
          const 规格 = row[2] ? String(row[2]).trim() : undefined;
          const 单位 = String(row[3] || '').trim();
          const 单价 = this.parseNumber(row[4]);
          const 含量 = this.parseNumber(row[5]);

          if (/^\d+[A-Z]-\d+$/.test(编号) || 编号.length > 0) {
            materials.push({
              编号: 编号 as any,
              名称,
              规格,
              单位,
              单价,
              含量,
              主材标记: false,
              材料号: undefined,
              材料类别: 3, // 默认为其他
              是否有明细: false
            });
          }
        }
      }
    }

    return materials;
  }

  // 解析数字的辅助方法
  private parseNumber(value: any): number {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[^\d.-]/g, '');
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    }
    return 0;
  }

  // 扫描整个工作表获取所有含量表数据
  scanAllMaterialData(): any[] {
    const materialList: any[] = [];
    let debugCount = 0;

    // 方法1: 基于"人 材 机 名 称"表头提取
    materialList.push(...this.extractMaterialsByHeader());

    // 方法2: 基于定额编号周围的数据提取
    materialList.push(...this.extractMaterialsAroundQuotaCodes());

    // 方法3: 基于边框区域提取
    materialList.push(...this.extractMaterialsFromBorderedRegions());

    // 去重处理
    const uniqueMaterials = this.deduplicateMaterials(materialList);

    console.log(`DEBUG: 总共找到 ${uniqueMaterials.length} 条材料数据`);
    return uniqueMaterials;
  }

  // 方法1: 基于"人 材 机 名 称"表头提取
  private extractMaterialsByHeader(): any[] {
    const materialList: any[] = [];
    let debugCount = 0;

    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      const rowData: string[] = [];

      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        rowData.push(String(value || ''));
      }

      const rowText = rowData.join(' ');

      // 查找包含"人 材 机 名 称"的行，这通常是材料表的开始
      if (rowText.includes('人 材 机 名 称') || rowText.includes('人材机名称') ||
          rowText.includes('人工') && rowText.includes('材料') && rowText.includes('机械')) {
        console.log(`DEBUG: 找到材料表头行 ${row}: ${rowText.substring(0, 100)}...`);

        // 从下一行开始提取材料数据，直到遇到下一个表头或空行
        for (let dataRow = row + 1; dataRow <= Math.min(row + 100, this.worksheet.rowCount); dataRow++) {
          const dataRowData: string[] = [];

          for (let col = 1; col <= Math.min(25, this.worksheet.columnCount); col++) {
            const cellInfo = this.getCellInfo(dataRow, col);
            const value = this.processExcelValue(cellInfo.value);
            dataRowData.push(String(value || ''));
          }

          const dataRowText = dataRowData.join(' ');

          // 如果遇到下一个表头或空行，停止提取
          if (dataRowText.includes('人 材 机 名 称') || dataRowText.includes('人材机名称') ||
              dataRowText.trim() === '' || dataRowText.includes('工作内容') ||
              dataRowText.includes('注 :') || dataRowText.includes('子目')) {
            break;
          }

          // 检查是否是有效的材料数据行
          if (dataRowData.length >= 2 && dataRowData[0] && dataRowData[1]) {
            const category = String(dataRowData[0]).trim();
            const materialName = String(dataRowData[1]).trim();

            // 更宽松的材料数据识别条件
            if (category && materialName && materialName.length > 1 && materialName.length < 100 &&
                (category === '人工' || category === '材料' || category === '机械' ||
                 category === '综合用工' || category === '载重汽车' || category === '船舶' ||
                 this.isValidMaterialCategory(category)) &&
                // 过滤掉子目名称和其他非材料内容
                !this.isSubItemName(materialName) &&
                !this.isSectionHeader(materialName) &&
                !this.isUnitHeader(materialName)) {

              debugCount++;
              console.log(`DEBUG: 找到材料数据 ${dataRow}: ${category} - ${materialName}`);

              // 向上查找最近的定额编号
              let nearestCode = this.findNearestQuotaCode(dataRow);

              // 提取更完整的材料信息
              const materialInfo = this.extractMaterialInfo(dataRowData, category);

              // 处理多行材料名称（按换行符分割）
              const materialNames = materialName.split(/\\n|\\r\\n|\\r/).map(name => name.trim()).filter(name => name);

              for (const singleMaterialName of materialNames) {
                if (singleMaterialName && singleMaterialName.length > 1 && singleMaterialName.length < 50) {
                  // 确定材料类别
                  let 材料类别 = this.determineMaterialCategory(category);

                  // 确定主材标记
                  const 主材标记 = 材料类别 === 2 && this.isMainMaterial(singleMaterialName);

                  materialList.push({
                    编号: nearestCode || '',
                    名称: singleMaterialName,
                    规格: materialInfo.规格,
                    单位: materialInfo.单位,
                    单价: materialInfo.单价,
                    含量: materialInfo.含量,
                    主材标记: 主材标记,
                    材料号: materialInfo.材料号,
                    材料类别: 材料类别,
                    是否有明细: materialInfo.是否有明细
                  });
                }
              }
            }
          }
        }
      }
    }

    console.log(`DEBUG: 方法1找到 ${debugCount} 条材料数据`);
    return materialList;
  }

  // 方法2: 基于定额编号周围的数据提取
  private extractMaterialsAroundQuotaCodes(): any[] {
    const materialList: any[] = [];
    let debugCount = 0;

    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      // 查找定额编号
      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(row, col);
        const value = this.processExcelValue(cellInfo.value);
        const str = String(value || '').trim();

        if (/^\d+[A-Z]-\d+$/.test(str)) {
          // 在定额编号周围查找材料数据
          for (let searchRow = Math.max(1, row - 10); searchRow <= Math.min(this.worksheet.rowCount, row + 20); searchRow++) {
            const searchRowData: string[] = [];

            for (let searchCol = 1; searchCol <= Math.min(25, this.worksheet.columnCount); searchCol++) {
              const searchCellInfo = this.getCellInfo(searchRow, searchCol);
              const searchValue = this.processExcelValue(searchCellInfo.value);
              searchRowData.push(String(searchValue || ''));
            }

            // 检查是否是材料数据行
            if (searchRowData.length >= 3 && searchRowData[0] && searchRowData[1]) {
              const category = String(searchRowData[0]).trim();
              const materialName = String(searchRowData[1]).trim();

              if (category && materialName && materialName.length > 1 && materialName.length < 100 &&
                  (category === '人工' || category === '材料' || category === '机械' ||
                   category === '综合用工' || category === '载重汽车' || category === '船舶' ||
                   this.isValidMaterialCategory(category)) &&
                  // 过滤掉子目名称和其他非材料内容
                  !this.isSubItemName(materialName) &&
                  !this.isSectionHeader(materialName) &&
                  !this.isUnitHeader(materialName) &&
                  !this.isTableHeader(materialName) &&
                  this.isValidMaterial(materialName)) {

                debugCount++;
                console.log(`DEBUG: 定额周围找到材料数据 ${searchRow}: ${category} - ${materialName}`);

                // 提取材料信息
                const materialInfo = this.extractMaterialInfo(searchRowData, category);

                // 确定材料类别
                let 材料类别 = this.determineMaterialCategory(category);

                // 确定主材标记
                const 主材标记 = 材料类别 === 2 && this.isMainMaterial(materialName);

                materialList.push({
                  编号: str,
                  名称: materialName,
                  规格: materialInfo.规格,
                  单位: materialInfo.单位,
                  单价: materialInfo.单价,
                  含量: materialInfo.含量,
                  主材标记: 主材标记,
                  材料号: materialInfo.材料号,
                  材料类别: 材料类别,
                  是否有明细: materialInfo.是否有明细
                });
              }
            }
          }
          break; // 找到一个定额编号后继续下一行
        }
      }
    }

    console.log(`DEBUG: 方法2找到 ${debugCount} 条材料数据`);
    return materialList;
  }

  // 方法3: 基于边框区域提取
  private extractMaterialsFromBorderedRegions(): any[] {
    const materialList: any[] = [];
    let debugCount = 0;

    // 扫描有边框的区域
    for (let row = 1; row <= this.worksheet.rowCount; row++) {
      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(row, col);

        // 如果是有边框的单元格，检查是否是材料表
        if (this.hasBorder(cellInfo)) {
          const regionData = this.getRegionDataAroundCell(row, col);

          // 检查区域是否包含材料数据
          for (const dataRow of regionData) {
            if (dataRow.length >= 3 && dataRow[0] && dataRow[1]) {
              const category = String(dataRow[0]).trim();
              const materialName = String(dataRow[1]).trim();

              if (category && materialName && materialName.length > 1 && materialName.length < 100 &&
                  (category === '人工' || category === '材料' || category === '机械' ||
                   category === '综合用工' || category === '载重汽车' || category === '船舶' ||
                   this.isValidMaterialCategory(category)) &&
                  // 过滤掉子目名称和其他非材料内容
                  !this.isSubItemName(materialName) &&
                  !this.isSectionHeader(materialName) &&
                  !this.isUnitHeader(materialName) &&
                  !this.isTableHeader(materialName) &&
                  this.isValidMaterial(materialName)) {

                debugCount++;
                console.log(`DEBUG: 边框区域找到材料数据: ${category} - ${materialName}`);

                // 提取材料信息
                const materialInfo = this.extractMaterialInfo(dataRow, category);

                // 确定材料类别
                let 材料类别 = this.determineMaterialCategory(category);

                // 确定主材标记
                const 主材标记 = 材料类别 === 2 && this.isMainMaterial(materialName);

                materialList.push({
                  编号: '',
                  名称: materialName,
                  规格: materialInfo.规格,
                  单位: materialInfo.单位,
                  单价: materialInfo.单价,
                  含量: materialInfo.含量,
                  主材标记: 主材标记,
                  材料号: materialInfo.材料号,
                  材料类别: 材料类别,
                  是否有明细: materialInfo.是否有明细
                });
              }
            }
          }
        }
      }
    }

    console.log(`DEBUG: 方法3找到 ${debugCount} 条材料数据`);
    return materialList;
  }

  // 辅助方法：查找最近的定额编号
  private findNearestQuotaCode(row: number): string {
    for (let searchRow = row - 1; searchRow >= Math.max(1, row - 20); searchRow--) {
      for (let col = 1; col <= this.worksheet.columnCount; col++) {
        const cellInfo = this.getCellInfo(searchRow, col);
        const value = this.processExcelValue(cellInfo.value);
        const str = String(value || '').trim();

        if (/^\d+[A-Z]-\d+$/.test(str)) {
          return str;
        }
      }
    }
    return '';
  }

  // 辅助方法：判断是否为有效的材料类别
  private isValidMaterialCategory(category: string): boolean {
    const validCategories = ['人工', '材料', '机械', '综合用工', '载重汽车', '船舶', '工日', '台班'];
    return validCategories.includes(category) || category.length <= 10;
  }

  // 辅助方法：确定材料类别
  private determineMaterialCategory(category: string): number {
    if (category === '人工' || category === '工日') {
      return 1; // 人工
    } else if (category === '机械' || category === '台班') {
      return 3; // 机械
    } else {
      return 2; // 材料
    }
  }

  // 辅助方法：获取单元格周围区域的数据
  private getRegionDataAroundCell(row: number, col: number): string[][] {
    const regionData: string[][] = [];

    for (let r = Math.max(1, row - 5); r <= Math.min(this.worksheet.rowCount, row + 10); r++) {
      const rowData: string[] = [];

      for (let c = Math.max(1, col - 5); c <= Math.min(this.worksheet.columnCount, col + 10); c++) {
        const cellInfo = this.getCellInfo(r, c);
        const value = this.processExcelValue(cellInfo.value);
        rowData.push(String(value || ''));
      }

      if (rowData.some(cell => cell.trim() !== '')) {
        regionData.push(rowData);
      }
    }

    return regionData;
  }

  // 辅助方法：去重材料数据
  private deduplicateMaterials(materials: any[]): any[] {
    const seen = new Set<string>();
    const uniqueMaterials: any[] = [];

    for (const material of materials) {
      const key = `${material.编号}-${material.名称}-${material.规格}`;
      if (!seen.has(key)) {
        seen.add(key);
        uniqueMaterials.push(material);
      }
    }

    return uniqueMaterials;
  }

  // 提取材料详细信息的辅助方法
  private extractMaterialInfo(dataRowData: string[], category: string): any {
    const materialInfo = {
      规格: '',
      单位: '',
      单价: 0,
      含量: 0,
      材料号: '',
      是否有明细: false
    };

    // 根据类别确定默认单位
    if (category === '人工') {
      materialInfo.单位 = '工日';
    } else if (category === '机械') {
      materialInfo.单位 = '台班';
    }

    // 扫描后续列寻找规格、单位、单价、含量等信息
    for (let col = 2; col < Math.min(20, dataRowData.length); col++) {
      const cellValue = String(dataRowData[col] || '').trim();

      if (!cellValue) continue;

      // 查找规格（通常包含特殊字符或数字+单位）
      if (!materialInfo.规格 && (cellValue.includes('Φ') || cellValue.includes('×') ||
          cellValue.includes('～') || cellValue.includes('-') || cellValue.includes('~') ||
          cellValue.includes('*') || cellValue.includes('#') ||
          /\d+[A-Za-z]/.test(cellValue) || /\d+[××]/.test(cellValue))) {
        materialInfo.规格 = cellValue;
      }

      // 查找单位
      if (!materialInfo.单位 && this.isUnit(cellValue)) {
        materialInfo.单位 = cellValue;
      }

      // 查找单价（数字，可能有小数点）
      if (materialInfo.单价 === 0 && /^\d+\.?\d*$/.test(cellValue)) {
        const numValue = this.parseNumber(cellValue);
        if (numValue > 0 && numValue < 1000000) { // 合理的单价范围
          materialInfo.单价 = numValue;
        }
      }

      // 查找含量（数字，可能有小数点）
      if (materialInfo.含量 === 0 && /^\d+\.?\d*$/.test(cellValue)) {
        const numValue = this.parseNumber(cellValue);
        if (numValue > 0 && numValue < 10000) { // 合理的含量范围
          materialInfo.含量 = numValue;
        }
      }

      // 查找材料号
      if (!materialInfo.材料号 && /^[A-Za-z0-9]+$/.test(cellValue) && cellValue.length <= 10) {
        materialInfo.材料号 = cellValue;
      }
    }

    // 如果仍然没有找到单位，尝试从材料名称中推断
    if (!materialInfo.单位) {
      const materialName = String(dataRowData[1] || '').trim();
      if (materialName.includes('kg') || materialName.includes('千克')) {
        materialInfo.单位 = 'kg';
      } else if (materialName.includes('m3') || materialName.includes('立方米')) {
        materialInfo.单位 = 'm3';
      } else if (materialName.includes('m2') || materialName.includes('平方米')) {
        materialInfo.单位 = 'm2';
      } else if (materialName.includes('m') || materialName.includes('米')) {
        materialInfo.单位 = 'm';
      } else if (materialName.includes('个') || materialName.includes('件')) {
        materialInfo.单位 = '个';
      } else if (materialName.includes('套')) {
        materialInfo.单位 = '套';
      } else if (materialName.includes('台')) {
        materialInfo.单位 = '台';
      } else if (materialName.includes('根')) {
        materialInfo.单位 = '根';
      } else if (materialName.includes('条')) {
        materialInfo.单位 = '条';
      } else if (materialName.includes('块')) {
        materialInfo.单位 = '块';
      } else if (materialName.includes('组')) {
        materialInfo.单位 = '组';
      } else if (materialName.includes('处')) {
        materialInfo.单位 = '处';
      } else if (materialName.includes('km')) {
        materialInfo.单位 = 'km';
      } else if (materialName.includes('头')) {
        materialInfo.单位 = '头';
      } else if (materialName.includes('t') || materialName.includes('吨')) {
        materialInfo.单位 = 't';
      } else if (materialName.includes('L') || materialName.includes('升')) {
        materialInfo.单位 = 'L';
      }
    }

    return materialInfo;
  }

  // 判断是否为单位的辅助方法
  private isUnit(value: string): boolean {
    const units = ['工日', '台班', 'kg', 'm3', 'm', '个', '套', '台', '根', '条', '块', '组', '处', 'km', '头', 't', 'L', 'm2',
                   '千克', '立方米', '平方米', '米', '件', '吨', '升', '公里', '千克', '公斤'];
    return units.includes(value);
  }

  // 判断是否为主材的辅助方法
  private isMainMaterial(materialName: string): boolean {
    const mainMaterials = ['钢材', '水泥', '木材', '砖', '砂', '石', '钢筋', '混凝土', '沥青', '管材', '电缆', '电线'];
    return mainMaterials.some(material => materialName.includes(material));
  }

  // 更智能的名称提取：在定额编号周围搜索名称
  private findQuotaName(row: number, col: number): string {
    // 首先尝试从同一行的下一列获取名称
    if (col + 1 <= this.worksheet.columnCount) {
      const nextCellInfo = this.getCellInfo(row, col + 1);
      const nextValue = this.processExcelValue(nextCellInfo.value);
      if (typeof nextValue === 'string') {
        const str = nextValue.trim();
        if (str.length > 3 && str.length < 100 && !/^\d+[A-Z]-\d+$/.test(str) &&
            !this.isSubItemName(str) && !this.isSectionHeader(str) && !this.isTableHeader(str)) {
          return str;
        }
      }
    }

    // 然后尝试从同一行的前一列获取名称
    if (col - 1 >= 1) {
      const prevCellInfo = this.getCellInfo(row, col - 1);
      const prevValue = this.processExcelValue(prevCellInfo.value);
      if (typeof prevValue === 'string') {
        const str = prevValue.trim();
        if (str.length > 3 && str.length < 100 && !/^\d+[A-Z]-\d+$/.test(str) &&
            !this.isSubItemName(str) && !this.isSectionHeader(str) && !this.isTableHeader(str)) {
          return str;
        }
      }
    }

    // 最后在周围区域搜索
    const startRow = Math.max(1, row - 3);
    const endRow = Math.min(this.worksheet.rowCount, row + 3);
    const startCol = Math.max(1, col - 3);
    const endCol = Math.min(this.worksheet.columnCount, col + 3);

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (r === row && c === col) continue; // 跳过定额编号本身

        const cellInfo = this.getCellInfo(r, c);
        const value = this.processExcelValue(cellInfo.value);
        if (typeof value === 'string') {
          const str = value.trim();
          if (str.length > 3 && str.length < 100 && !/^\d+[A-Z]-\d+$/.test(str) &&
              !this.isSubItemName(str) && !this.isSectionHeader(str) &&
              !this.isUnitHeader(str) && !this.isTableHeader(str)) {
            return str;
          }
        }
      }
    }
    return '';
  }

  // 判断是否为子目名称的辅助方法
  private isSubItemName(name: string): boolean {
    const subItemNames = ['工作内容', '附注', '子目', '编号', '名称', '单位', '基价', '人工', '材料', '机械', '管理费', '利润', '其他', '规格', '单价', '含量', '材料号', '是否有明细', '子目编号', '子目名称'];
    return subItemNames.some(subName => name.includes(subName));
  }

  // 判断是否为章节标题的辅助方法
  private isSectionHeader(name: string): boolean {
    const sectionHeaders = ['第一章', '第二章', '第三章', '第四章', '第五章', '第六章', '第七章', '第八章', '第九章', '第十章', '第一节', '第二节', '第三节', '第四节', '第五节', '第六节', '第七节', '第八节', '第九节', '第十节'];
    return sectionHeaders.some(header => name.includes(header));
  }

  // 判断是否为单位表头的辅助方法
  private isUnitHeader(name: string): boolean {
    const unitHeaders = ['工日', '台班', 'kg', 'm3', 'm', '个', '套', '台', '根', '条', '块', '组', '处', 'km', '头', 't', 'L', 'm2', '千克', '立方米', '平方米', '米', '件', '吨', '升', '公里', '千克', '公斤'];
    return unitHeaders.some(header => name.includes(header));
  }

  // 判断是否为表格表头的辅助方法
  private isTableHeader(name: string): boolean {
    const tableHeaders = ['人 材 机 名 称', '人材机名称', '工作内容', '附注', '子目', '编号', '名称', '单位', '基价', '人工', '材料', '机械', '管理费', '利润', '其他', '规格', '单价', '含量', '材料号', '是否有明细', '消耗量', '工程量计算规则', '减振装置安装', '燃气采暖炉', '燃气开水炉', '手动放风阀', '单气嘴'];
    return tableHeaders.some(header => name.includes(header));
  }

  // 判断是否为有效材料的辅助方法
  private isValidMaterial(materialName: string): boolean {
    const validMaterials = ['钢材', '水泥', '木材', '砖', '砂', '石', '钢筋', '混凝土', '沥青', '管材', '电缆', '电线', '工日', '台班', 'kg', 'm3', 'm', '个', '套', '台', '根', '条', '块', '组', '处', 'km', '头', 't', 'L', 'm2', '千克', '立方米', '平方米', '米', '件', '吨', '升', '公里', '千克', '公斤'];
    return validMaterials.some(material => materialName.includes(material));
  }
}
