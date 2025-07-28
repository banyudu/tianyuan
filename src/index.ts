import * as path from 'path';
import { ExcelAnalyzer } from './excel-analyzer';
import { DataExtractor } from './data-extractor';
import { FileExporter } from './file-exporter';

async function main(): Promise<void> {
  try {
    console.log('=== Excel转换器启动 ===');

    // 文件路径配置
    const inputFile = path.join(__dirname, '../sample/input.xlsx');
    const outputDir = path.join(__dirname, '../output');

    console.log(`输入文件: ${inputFile}`);
    console.log(`输出目录: ${outputDir}`);

    // 步骤1: 分析Excel文件
    console.log('\n步骤1: 加载并分析Excel文件...');
    const analyzer = new ExcelAnalyzer();
    await analyzer.loadFile(inputFile);

    // 步骤2: 基于定额编号识别数据区域
    console.log('\n步骤2: 基于定额编号识别数据区域...');
    const regions = analyzer.scanDefineCodeRegions();

    console.log(`\n找到 ${regions.length} 个数据区域`);

    // 步骤3: 直接提取工作内容
    console.log('\n步骤3: 提取工作内容数据...');
    const workContentData: {编号: string, 工作内容: string}[] = [];

    for (const region of regions) {
      if (region.metadata?.type === '工作内容' && region.metadata.workContentRow > 0) {
        const workContent = analyzer.getWorkContentFromRow(region.metadata.workContentRow, region.metadata.codes);
        if (workContent) {
          workContentData.push(workContent);
          console.log(`提取工作内容: ${workContent.编号} -> ${workContent.工作内容.substring(0, 50)}...`);
        }
      }
    }

    // 步骤4: 提取含量表数据
    console.log('\n步骤4: 提取含量表数据...');
    const materialData = analyzer.scanAllMaterialData();
    console.log(`提取含量表: ${materialData.length} 条记录`);

    // 步骤5: 提取附注信息数据
    console.log('\n步骤5: 提取附注信息数据...');
    const noteInfoData = analyzer.scanAllNoteInfo();
    console.log(`提取附注信息: ${noteInfoData.length} 条记录`);

    // 步骤6: 提取子目信息数据
    console.log('\n步骤6: 提取子目信息数据...');
    const subItemData = analyzer.scanAllSubItemInfo();
    console.log(`提取子目信息: ${subItemData.length} 条记录`);

    // 步骤7: 显示提取结果
    console.log('\n步骤7: 数据提取完成');
    console.log(`工作内容: ${workContentData.length} 条记录`);
    console.log(`含量表: ${materialData.length} 条记录`);
    console.log(`附注信息: ${noteInfoData.length} 条记录`);
    console.log(`子目信息: ${subItemData.length} 条记录`);

    const totalRecords = workContentData.length + materialData.length + noteInfoData.length + subItemData.length;
    console.log(`总计: ${totalRecords} 条记录`);

    if (totalRecords === 0) {
      console.warn('警告: 没有提取到任何数据记录！');
    }

    // 步骤8: 导出文件
    console.log('\n步骤8: 导出文件...');

    // 构建输出数据结构
    const extractedData = {
      附注信息: noteInfoData,
      工作内容: workContentData,
      子目信息: subItemData,
      含量表: materialData,
    };

    // 导出CSV文件（用于调试和验证）
    const exporter = new FileExporter();
    exporter.printDataSummary(extractedData);

    // 导出CSV文件
    await exporter.exportToCSV(extractedData, 'output/csv');

    // 导出Excel文件
    await exporter.exportToExcel(extractedData, 'output/excel');

    console.log('\n=== 转换完成 ===');
    console.log('请检查输出文件是否符合预期格式');
    console.log('\n数据统计:');
    console.log(`- 附注信息: ${extractedData.附注信息.length} 条`);
    console.log(`- 工作内容: ${extractedData.工作内容.length} 条`);
    console.log(`- 子目信息: ${extractedData.子目信息.length} 条`);
    console.log(`- 含量表: ${extractedData.含量表.length} 条`);

    // 显示一些样本数据
    if (extractedData.工作内容.length > 0) {
      console.log('\n工作内容样本:');
      extractedData.工作内容.slice(0, 3).forEach(item => {
        console.log(`  ${item.编号}: ${item.工作内容}`);
      });
    }

  } catch (error) {
    console.error('转换过程中发生错误:', error);
    process.exit(1);
  }
}

// 如果直接运行此文件，执行main函数
if (require.main === module) {
  main().catch(console.error);
}

export { main };
