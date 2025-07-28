import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { 附注信息行, 工作内容行, 子目信息表行, 材料表行 } from './types';

export class FileExporter {
  // 导出为CSV文件
  async exportToCSV(
    data: {
      附注信息: 附注信息行[];
      工作内容: 工作内容行[];
      子目信息: 子目信息表行[];
      含量表: 材料表行[];
    },
    outputDir: string
  ): Promise<void> {
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // 导出附注信息
    const fuZhuCSV = this.convertToCSV(data.附注信息, ['编号', '附注信息']);
    fs.writeFileSync(path.join(outputDir, '附注信息.csv'), fuZhuCSV, 'utf8');

    // 导出工作内容
    const gongZuoCSV = this.convertToCSV(data.工作内容, ['编号', '工作内容']);
    fs.writeFileSync(path.join(outputDir, '工作内容.csv'), gongZuoCSV, 'utf8');

    // 导出子目信息
    const ziMuCSV = this.convertToCSV(
      data.子目信息,
      [
        '',
        '定额号',
        '子目名称',
        '基价',
        '人工',
        '材料',
        '机械',
        '管理费',
        '利润',
        '其他',
        '图片名称',
        '',
      ],
      item => [
        item.symbol,
        '', // 定额号列（在子目信息中没有单独的定额号）
        item.子目名称,
        item.基价,
        item.人工,
        item.材料,
        item.机械,
        item.管理费,
        item.利润,
        item.其他,
        item.图片名称 || '',
        '',
      ]
    );
    fs.writeFileSync(path.join(outputDir, '子目信息.csv'), ziMuCSV, 'utf8');

    // 导出含量表
    const hanLiangCSV = this.convertToCSV(
      data.含量表,
      [
        '编号',
        '名称',
        '规格',
        '单位',
        '单价',
        '含量',
        '主材标记',
        '材料号',
        '材料类别',
        '是否有明细',
        '',
        '',
        '',
      ],
      item => [
        item.编号,
        item.名称,
        item.规格 || '',
        item.单位,
        item.单价,
        item.含量,
        item.主材标记 ? '*' : '',
        item.材料号 || '',
        item.材料类别,
        item.是否有明细 ? '是' : '',
        '',
        '',
        '',
      ]
    );
    fs.writeFileSync(path.join(outputDir, '含量表.csv'), hanLiangCSV, 'utf8');

    console.log(`CSV文件已导出到: ${outputDir}`);
  }

  // 导出为Excel文件
  async exportToExcel(
    data: {
      附注信息: 附注信息行[];
      工作内容: 工作内容行[];
      子目信息: 子目信息表行[];
      含量表: 材料表行[];
    },
    outputDir: string
  ): Promise<void> {
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // 创建工作内容和附注信息的合并文件
    const workbook1 = new ExcelJS.Workbook();

    // 工作内容工作表
    const workSheet = workbook1.addWorksheet('工作内容');
    workSheet.addRow(['编号', '工作内容']);
    data.工作内容.forEach(item => {
      workSheet.addRow([item.编号, item.工作内容]);
    });

    // 附注信息工作表
    const noteSheet = workbook1.addWorksheet('附注信息');
    noteSheet.addRow(['编号', '附注信息']);
    data.附注信息.forEach(item => {
      noteSheet.addRow([item.编号, item.附注信息]);
    });

    await workbook1.xlsx.writeFile(path.join(outputDir, '工作内容_附注信息.xlsx'));

    // 创建子目信息文件
    const workbook2 = new ExcelJS.Workbook();
    const zimuSheet = workbook2.addWorksheet('子目信息');
    zimuSheet.addRow([
      '',
      '定额号',
      '子目名称',
      '基价',
      '人工',
      '材料',
      '机械',
      '管理费',
      '利润',
      '其他',
      '图片名称',
      '',
    ]);

    data.子目信息.forEach(item => {
      zimuSheet.addRow([
        item.symbol,
        '', // 定额号
        item.子目名称,
        item.基价,
        item.人工,
        item.材料,
        item.机械,
        item.管理费,
        item.利润,
        item.其他,
        item.图片名称 || '',
        '',
      ]);
    });

    await workbook2.xlsx.writeFile(path.join(outputDir, '子目信息.xlsx'));

    // 创建含量表文件
    const workbook3 = new ExcelJS.Workbook();
    const hanliangSheet = workbook3.addWorksheet('含量表');
    hanliangSheet.addRow([
      '编号',
      '名称',
      '规格',
      '单位',
      '单价',
      '含量',
      '主材标记',
      '材料号',
      '材料类别',
      '是否有明细',
      '',
      '',
      '',
    ]);

    data.含量表.forEach(item => {
      hanliangSheet.addRow([
        item.编号,
        item.名称,
        item.规格 || '',
        item.单位,
        item.单价,
        item.含量,
        item.主材标记 ? '*' : '',
        item.材料号 || '',
        item.材料类别,
        item.是否有明细 ? '是' : '',
        '',
        '',
        '',
      ]);
    });

    await workbook3.xlsx.writeFile(path.join(outputDir, '含量表.xlsx'));

    console.log(`Excel文件已导出到: ${outputDir}`);
  }

  // 转换为CSV格式的通用方法
  private convertToCSV<T>(data: T[], headers: string[], mapper?: (item: T) => any[]): string {
    const rows: string[] = [];

    // 添加标题行
    rows.push(headers.map(h => this.escapeCSV(h)).join(','));

    // 添加数据行
    data.forEach(item => {
      let values: any[];
      if (mapper) {
        values = mapper(item);
      } else {
        values = Object.values(item);
      }

      const csvRow = values.map(v => this.escapeCSV(String(v || ''))).join(',');
      rows.push(csvRow);
    });

    return rows.join('\n');
  }

  // CSV字段转义
  private escapeCSV(value: string): string {
    if (value.includes(',') || value.includes('"') || value.includes('\n')) {
      return `"${value.replace(/"/g, '""')}"`;
    }
    return value;
  }

  // 打印数据统计信息
  printDataSummary(data: {
    附注信息: 附注信息行[];
    工作内容: 工作内容行[];
    子目信息: 子目信息表行[];
    含量表: 材料表行[];
  }): void {
    console.log('\n=== 数据提取结果统计 ===');
    console.log(`附注信息: ${data.附注信息.length} 条`);
    console.log(`工作内容: ${data.工作内容.length} 条`);
    console.log(`子目信息: ${data.子目信息.length} 条`);
    console.log(`含量表: ${data.含量表.length} 条`);

    console.log('\n=== 示例数据 ===');
    if (data.附注信息.length > 0) {
      console.log('附注信息示例:', data.附注信息[0]);
    }
    if (data.工作内容.length > 0) {
      console.log('工作内容示例:', data.工作内容[0]);
    }
    if (data.子目信息.length > 0) {
      console.log('子目信息示例:', data.子目信息[0]);
    }
    if (data.含量表.length > 0) {
      console.log('含量表示例:', data.含量表[0]);
    }
  }
}
