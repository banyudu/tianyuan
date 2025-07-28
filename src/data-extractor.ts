import { 子目编号, 附注信息行, 工作内容行, 子目信息表行, 材料表行, 材料类别 } from './types';
import { TableRegion, ExcelAnalyzer } from './excel-analyzer';

export class DataExtractor {
  // 提取附注信息
  extractFuZhuInfo(region: TableRegion, analyzer: ExcelAnalyzer): 附注信息行[] {
    const data = analyzer.getRegionData(region);
    const result: 附注信息行[] = [];

    // 寻找标题行
    let startRow = 0;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      if (rowText.includes('编号') && rowText.includes('附注')) {
        startRow = i + 1;
        break;
      }
    }

    // 从标题行后开始处理数据
    for (let i = startRow; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 2) {
        const 编号 = this.cleanString(row[0]);
        const 附注信息 = this.cleanString(row[1]);

        // 检查是否是有效的编号格式（逗号分隔的定额号）
        if (编号 && 附注信息 && this.isValidDefineCode(编号)) {
          result.push({ 编号, 附注信息 });
        }
      }
    }

    return result;
  }

  // 提取工作内容
  extractGongZuoNeiRong(region: TableRegion, analyzer: ExcelAnalyzer): 工作内容行[] {
    const data = analyzer.getRegionData(region);
    const result: 工作内容行[] = [];

    // 寻找标题行
    let startRow = 0;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      if (rowText.includes('编号') && rowText.includes('工作内容')) {
        startRow = i + 1;
        break;
      }
    }

    // 从标题行后开始处理数据
    for (let i = startRow; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 2) {
        const 编号 = this.cleanString(row[0]);
        const 工作内容 = this.cleanString(row[1]);

        // 检查是否是有效的编号格式（逗号分隔的定额号）
        if (编号 && 工作内容 && this.isValidDefineCode(编号)) {
          result.push({ 编号, 工作内容 });
        }
      }
    }

    return result;
  }

  // 提取子目信息
  extractZiMuInfo(region: TableRegion, analyzer: ExcelAnalyzer): 子目信息表行[] {
    const data = analyzer.getRegionData(region);
    const result: 子目信息表行[] = [];

    // 寻找标题行
    let startRow = 0;
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      if (rowText.includes('子目名称') || (rowText.includes('定额') && rowText.includes('名称'))) {
        startRow = i + 1;
        break;
      }
    }

    // 从标题行后开始处理数据
    for (let i = startRow; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 3) {
        // 识别层级符号
        let symbol = '';
        let 子目名称 = '';

        // 检查第一列是否包含层级符号
        const firstCol = this.cleanString(row[0]);
        const secondCol = this.cleanString(row[1]);

        if (firstCol.includes('$')) {
          symbol = firstCol;
          子目名称 = secondCol;
        } else if (firstCol.includes('第') && firstCol.includes('章')) {
          symbol = '$';
          子目名称 = firstCol;
        } else if (firstCol.includes('第') && firstCol.includes('节')) {
          symbol = '$$';
          子目名称 = firstCol;
        } else if (firstCol.match(/[一二三四五六七八九十]+、/)) {
          symbol = '$$$';
          子目名称 = firstCol;
        } else if (firstCol.match(/\d+\./)) {
          symbol = '$$$$';
          子目名称 = firstCol;
        } else {
          // 可能是普通数据行
          子目名称 = secondCol || firstCol;
        }

        // 解析数字列（从第3列开始或者适当调整）
        const baseIndex = symbol ? 2 : 1;
        const 基价 = this.parseNumber(row[baseIndex + 1]);
        const 人工 = this.parseNumber(row[baseIndex + 2]);
        const 材料 = this.parseNumber(row[baseIndex + 3]);
        const 机械 = this.parseNumber(row[baseIndex + 4]);
        const 管理费 = this.parseNumber(row[baseIndex + 5]);
        const 利润 = this.parseNumber(row[baseIndex + 6]);
        const 其他 = this.parseNumber(row[baseIndex + 7]);
        const 图片名称 = row[baseIndex + 8] ? this.cleanString(row[baseIndex + 8]) : undefined;

        if (子目名称 && 子目名称.trim()) {
          result.push({
            symbol: symbol as any,
            子目名称,
            基价,
            人工,
            材料,
            机械,
            管理费,
            利润,
            其他,
            图片名称,
          });
        }
      }
    }

    return result;
  }

  // 提取含量表信息
  extractHanLiangBiao(region: TableRegion, analyzer: ExcelAnalyzer): 材料表行[] {
    const data = analyzer.getRegionData(region);
    const result: 材料表行[] = [];

    // 寻找标题行
    let startRow = 0;
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const rowText = data[i].join(' ').toLowerCase();
      if (rowText.includes('编号') && (rowText.includes('名称') || rowText.includes('材料')) && rowText.includes('单位')) {
        startRow = i + 1;
        break;
      }
    }

    // 从标题行后开始处理数据
    for (let i = startRow; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 6) {
        const 编号 = this.cleanString(row[0]) as 子目编号;
        const 名称 = this.cleanString(row[1]);
        const 规格 = row[2] ? this.cleanString(row[2]) : undefined;
        const 单位 = this.cleanString(row[3]);
        const 单价 = this.parseNumber(row[4]);
        const 含量 = this.parseNumber(row[5]);
        const 主材标记 = this.parseBoolean(row[6]);
        const 材料号 = row[7] ? this.cleanString(row[7]) : undefined;
        const 材料类别 = this.parseMaterialCategory(row[8]);
        const 是否有明细 = this.parseBoolean(row[9]);

        // 验证必要字段
        if (编号 && 名称 && 单位 && this.isValidMaterialCode(编号)) {
          result.push({
            编号,
            名称,
            规格,
            单位,
            单价,
            含量,
            主材标记,
            材料号,
            材料类别,
            是否有明细,
          });
        }
      }
    }

    return result;
  }

  // 工具方法：清理字符串
  private cleanString(value: any): string {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  // 工具方法：检查是否是有效的定额编号格式
  private isValidDefineCode(code: string): boolean {
    if (!code) return false;

    // 检查是否包含定额编号格式（如 2B-84,2B-85,2B-86）
    const parts = code.split(',');
    return parts.some(part => /^\d+[A-Z]-\d+$/.test(part.trim()));
  }

  // 工具方法：检查是否是有效的材料编号格式
  private isValidMaterialCode(code: string): boolean {
    if (!code) return false;

    // 检查是否是定额编号格式（如 1B-1）或其他有效格式
    return /^\d+[A-Z]-\d+$/.test(code) || code.length > 1;
  }

  // 工具方法：解析数字
  private parseNumber(value: any): number {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[^\d.-]/g, '');
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    }
    return 0;
  }

  // 工具方法：解析布尔值
  private parseBoolean(value: any): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') {
      const str = value.toLowerCase().trim();
      return str === 'true' || str === '1' || str === '是' || str === '*' || str === '√';
    }
    return false;
  }

  // 工具方法：解析材料类别
  private parseMaterialCategory(value: any): 材料类别 {
    if (typeof value === 'number') {
      switch (value) {
        case 1:
          return 材料类别.主材;
        case 2:
          return 材料类别.辅材;
        case 3:
          return 材料类别.其他;
        default:
          return 材料类别.其他;
      }
    }
    if (typeof value === 'string') {
      const str = value.trim();
      if (str === '1' || str === '主材') return 材料类别.主材;
      if (str === '2' || str === '辅材') return 材料类别.辅材;
      return 材料类别.其他;
    }
    return 材料类别.其他;
  }

  // 智能识别并提取所有数据
  extractAllData(
    regions: TableRegion[],
    analyzer: ExcelAnalyzer
  ): {
    附注信息: 附注信息行[];
    工作内容: 工作内容行[];
    子目信息: 子目信息表行[];
    含量表: 材料表行[];
  } {
    const result = {
      附注信息: [] as 附注信息行[],
      工作内容: [] as 工作内容行[],
      子目信息: [] as 子目信息表行[],
      含量表: [] as 材料表行[],
    };

    for (const region of regions) {
      const dataType = analyzer.identifyDataType(region);

      console.log(`处理区域: ${dataType} (${region.startRow}-${region.endRow})`);

      switch (dataType) {
        case '附注信息':
          const fuZhuData = this.extractFuZhuInfo(region, analyzer);
          console.log(`  提取附注信息 ${fuZhuData.length} 条`);
          result.附注信息.push(...fuZhuData);
          break;
        case '工作内容':
          const gongZuoData = this.extractGongZuoNeiRong(region, analyzer);
          console.log(`  提取工作内容 ${gongZuoData.length} 条`);
          result.工作内容.push(...gongZuoData);
          break;
        case '子目信息':
          const ziMuData = this.extractZiMuInfo(region, analyzer);
          console.log(`  提取子目信息 ${ziMuData.length} 条`);
          result.子目信息.push(...ziMuData);
          break;
        case '含量表':
          const hanLiangData = this.extractHanLiangBiao(region, analyzer);
          console.log(`  提取含量表 ${hanLiangData.length} 条`);
          result.含量表.push(...hanLiangData);
          break;
        default:
          console.log(`  未识别的数据类型: ${dataType}，尝试自动分析...`);
          // 尝试根据数据结构自动判断
          this.tryAutoExtract(region, analyzer, result);
          break;
      }
    }

    return result;
  }

  // 尝试自动提取数据
  private tryAutoExtract(
    region: TableRegion,
    analyzer: ExcelAnalyzer,
    result: {
      附注信息: 附注信息行[];
      工作内容: 工作内容行[];
      子目信息: 子目信息表行[];
      含量表: 材料表行[];
    }
  ): void {
    const data = analyzer.getRegionData(region);

    if (data.length === 0) return;

    // 分析数据特征
    const avgCols = data.reduce((sum, row) => sum + row.length, 0) / data.length;
    const hasDefinaCodes = data.some(row =>
      row.some(cell => {
        const str = String(cell || '');
        return /^\d+[A-Z]-\d+/.test(str);
      })
    );

    const hasCommaList = data.some(row =>
      row.some(cell => {
        const str = String(cell || '');
        return str.includes(',') && /\d+[A-Z]-\d+/.test(str);
      })
    );

    console.log(`  自动分析: 平均列数=${avgCols.toFixed(1)}, 有定额号=${hasDefinaCodes}, 有逗号列表=${hasCommaList}`);

    if (hasCommaList) {
      // 可能是工作内容或附注信息
      const gongZuoData = this.extractGongZuoNeiRong(region, analyzer);
      const fuZhuData = this.extractFuZhuInfo(region, analyzer);

      if (gongZuoData.length > fuZhuData.length) {
        console.log(`  自动识别为工作内容: ${gongZuoData.length} 条`);
        result.工作内容.push(...gongZuoData);
      } else if (fuZhuData.length > 0) {
        console.log(`  自动识别为附注信息: ${fuZhuData.length} 条`);
        result.附注信息.push(...fuZhuData);
      }
    } else if (avgCols >= 8 && hasDefinaCodes) {
      // 可能是含量表
      const hanLiangData = this.extractHanLiangBiao(region, analyzer);
      console.log(`  自动识别为含量表: ${hanLiangData.length} 条`);
      result.含量表.push(...hanLiangData);
    } else if (avgCols >= 5) {
      // 可能是子目信息
      const ziMuData = this.extractZiMuInfo(region, analyzer);
      console.log(`  自动识别为子目信息: ${ziMuData.length} 条`);
      result.子目信息.push(...ziMuData);
    }
  }
}
