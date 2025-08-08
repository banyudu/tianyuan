import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { StructuredDocument, Chapter, Section, SubSection, TableArea, NormInfo, ResourceConsumption } from './src/structured-types';

class FinalResultsExporter {
  private structuredData: StructuredDocument;

  constructor(structuredDataPath: string) {
    console.log(`Loading structured data from: ${structuredDataPath}`);
    const rawData = fs.readFileSync(structuredDataPath, 'utf8');
    this.structuredData = JSON.parse(rawData);
  }

  // Generate 含量表.csv - Material consumption table
  private generateHanLiangBiao(): string {
    console.log('Generating 含量表.csv...');
    const rows: string[] = [];

    // Header
    rows.push('编号,名称,规格,单位,单价,含量,主材标记,材料号,材料类别,是否有明细,,,');

    // Process all chapters, sections, and norms
    this.structuredData.chapters.forEach(chapter => {
      this.processChapterForHanLiangBiao(chapter, rows);
    });

    return rows.join('\n');
  }

  private processChapterForHanLiangBiao(chapter: Chapter, rows: string[]): void {
    // Process table areas directly in chapter
    chapter.tableAreas?.forEach(tableArea => {
      this.processTableAreaForHanLiangBiao(tableArea, rows);
    });

    // Process sections
    chapter.sections.forEach(section => {
      this.processSectionForHanLiangBiao(section, rows);
    });
  }

  private processSectionForHanLiangBiao(section: Section, rows: string[]): void {
    // Process table areas directly in section
    section.tableAreas?.forEach(tableArea => {
      this.processTableAreaForHanLiangBiao(tableArea, rows);
    });

    // Process subsections
    section.subSections.forEach(subSection => {
      this.processSubSectionForHanLiangBiao(subSection, rows);
    });
  }

  private processSubSectionForHanLiangBiao(subSection: SubSection, rows: string[]): void {
    // Process table areas in subsection
    subSection.tableAreas.forEach(tableArea => {
      this.processTableAreaForHanLiangBiao(tableArea, rows);
    });

    // Process children subsections
    subSection.children.forEach(child => {
      this.processSubSectionForHanLiangBiao(child, rows);
    });
  }

  private processTableAreaForHanLiangBiao(tableArea: TableArea, rows: string[]): void {
    // Process norms in this table area
    if (tableArea.norms) {
      tableArea.norms.forEach(norm => {
        if (norm.resources && norm.resources.length > 0) {
          norm.resources.forEach(resource => {
            // Format: 编号,名称,规格,单位,单价,含量,主材标记,材料号,材料类别,是否有明细
            const row = [
              norm.code,                                    // 编号
              this.escapeCsvField(resource.name),           // 名称
              this.escapeCsvField(resource.specification || ''), // 规格
              this.escapeCsvField(resource.unit),           // 单位
              '0',                                          // 单价 (placeholder)
              resource.consumption,                         // 含量
              resource.isPrimary ? '*' : '',               // 主材标记
              '',                                          // 材料号 (empty)
              resource.categoryCode.toString(),            // 材料类别
              '',                                          // 是否有明细 (empty)
              '',                                          // Empty column
              '',                                          // Empty column
              ''                                           // Empty column
            ];
            rows.push(row.join(','));
          });
        }
      });
    }
  }

  // Generate 子目信息.csv - Norm hierarchy information
  private generateZiMuXinXi(): string {
    console.log('Generating 子目信息.csv...');
    const rows: string[] = [];

    // Header
    rows.push(',定额号,子目名称,单位,基价,人工,材料,机械,图片名称,');

    let sequenceNumber = 1;

    // Process all chapters
    this.structuredData.chapters.forEach(chapter => {
      // Add chapter header
      rows.push([
        '$',
        '',
        this.escapeCsvField(chapter.name),
        '',
        '',
        '',
        '',
        '',
        ''
      ].join(','));

      this.processChapterForZiMuXinXi(chapter, rows, sequenceNumber);
      sequenceNumber = this.getNextSequenceNumber(rows);
    });

    return rows.join('\n');
  }

  private processChapterForZiMuXinXi(chapter: Chapter, rows: string[], startSeqNum: number): void {
    let sequenceNumber = startSeqNum;

    // Process sections
    chapter.sections.forEach(section => {
      // Add section header
      rows.push([
        '$$',
        '',
        this.escapeCsvField(section.name),
        '',
        '',
        '',
        '',
        '',
        ''
      ].join(','));

      this.processSectionForZiMuXinXi(section, rows, sequenceNumber);
      sequenceNumber = this.getNextSequenceNumber(rows);
    });
  }

  private processSectionForZiMuXinXi(section: Section, rows: string[], startSeqNum: number): void {
    let sequenceNumber = startSeqNum;

    // Process subsections
    section.subSections.forEach(subSection => {
      // Add subsection header
      rows.push([
        '$$$',
        '',
        this.escapeCsvField(subSection.name),
        '',
        '',
        '',
        '',
        '',
        ''
      ].join(','));

      this.processSubSectionForZiMuXinXi(subSection, rows, sequenceNumber);
      sequenceNumber = this.getNextSequenceNumber(rows);
    });

    // Process table areas directly in section
    section.tableAreas?.forEach(tableArea => {
      this.processTableAreaForZiMuXinXi(tableArea, rows, sequenceNumber);
      sequenceNumber = this.getNextSequenceNumber(rows);
    });
  }

  private processSubSectionForZiMuXinXi(subSection: SubSection, rows: string[], startSeqNum: number): void {
    let sequenceNumber = startSeqNum;

    // Process table areas in subsection
    subSection.tableAreas.forEach(tableArea => {
      this.processTableAreaForZiMuXinXi(tableArea, rows, sequenceNumber);
      sequenceNumber = this.getNextSequenceNumber(rows);
    });

    // Process children subsections
    subSection.children.forEach(child => {
      this.processSubSectionForZiMuXinXi(child, rows, sequenceNumber);
      sequenceNumber = this.getNextSequenceNumber(rows);
    });
  }

  private processTableAreaForZiMuXinXi(tableArea: TableArea, rows: string[], startSeqNum: number): void {
    // Process norms in this table area
    if (tableArea.norms) {
      tableArea.norms.forEach(norm => {
        // Get the full name from the structure if available
        const fullName = this.getNormFullName(norm, tableArea);

        // Calculate totals for 人工, 材料, 机械
        const totals = this.calculateNormTotals(norm);

        // Format: 符号,定额号,子目名称,基价,人工,材料,机械,管理费,利润,其他,图片名称,
        const row = [
          '',                                     // 符号 (norm level)
          norm.code,                                  // 定额号
          this.escapeCsvField(fullName),              // 子目名称
          0, // 基价
          0, // 人工
          0, // 材料
          0, // 机械
          0, // 管理费
          0, // 利润
          0, // 其他
          '' // 图片名称 (empty)
        ];
        rows.push(row.join(','));
      });
    }
  }

  private getNormFullName(norm: NormInfo, tableArea: TableArea): string {
    // Try to get full name from the normNamesRows structure
    if (tableArea.structure?.normNamesRows) {
      const normName = tableArea.structure.normNamesRows.normNames.find(n => n.normCode === norm.code);
      if (normName && normName.fullName) {
        return normName.fullName;
      }
    }
    // Fallback to just the norm code
    return norm.code;
  }

  private calculateNormTotals(norm: NormInfo): { labor: number; materials: number; machinery: number } {
    const totals = { labor: 0, materials: 0, machinery: 0 };

    if (norm.resources) {
      norm.resources.forEach(resource => {
        const consumption = parseFloat(resource.consumption) || 0;

        switch (resource.categoryCode) {
          case 1: // 人工
            totals.labor += consumption;
            break;
          case 2: // 材料
          case 5: // 主材
            totals.materials += consumption;
            break;
          case 3: // 机械
            totals.machinery += consumption;
            break;
        }
      });
    }

    return totals;
  }

  private getNextSequenceNumber(rows: string[]): number {
    // Parse the last row to get the sequence number and increment
    if (rows.length > 1) {
      const lastRow = rows[rows.length - 1];
      const seqNum = parseInt(lastRow.split(',')[0]);
      return seqNum + 1;
    }
    return 1;
  }

  // Generate 工作内容.csv - Work content information
  private generateGongZuoNeiRong(): string {
    console.log('Generating 工作内容.csv...');
    const rows: string[] = [];

    // Header
    rows.push('编号,工作内容');

    // Process all chapters, sections, and table areas to find work content
    this.structuredData.chapters.forEach(chapter => {
      this.processChapterForGongZuoNeiRong(chapter, rows);
    });

    return rows.join('\n');
  }

  private processChapterForGongZuoNeiRong(chapter: Chapter, rows: string[]): void {
    // Process table areas directly in chapter
    chapter.tableAreas?.forEach(tableArea => {
      this.processTableAreaForGongZuoNeiRong(tableArea, rows);
    });

    // Process sections
    chapter.sections.forEach(section => {
      this.processSectionForGongZuoNeiRong(section, rows);
    });
  }

  private processSectionForGongZuoNeiRong(section: Section, rows: string[]): void {
    // Process table areas directly in section
    section.tableAreas?.forEach(tableArea => {
      this.processTableAreaForGongZuoNeiRong(tableArea, rows);
    });

    // Process subsections
    section.subSections.forEach(subSection => {
      this.processSubSectionForGongZuoNeiRong(subSection, rows);
    });
  }

  private processSubSectionForGongZuoNeiRong(subSection: SubSection, rows: string[]): void {
    // Process table areas in subsection
    subSection.tableAreas.forEach(tableArea => {
      this.processTableAreaForGongZuoNeiRong(tableArea, rows);
    });

    // Process children subsections
    subSection.children.forEach(child => {
      this.processSubSectionForGongZuoNeiRong(child, rows);
    });
  }

  private processTableAreaForGongZuoNeiRong(tableArea: TableArea, rows: string[]): void {
    // Check if table area has work content
    if (tableArea.workContent && tableArea.workContent.trim()) {
      // Extract norm codes for this work content
      const normCodes = tableArea.normCodes.join(',');

      rows.push([
        `"${normCodes}"`,
        this.escapeCsvField(tableArea.workContent)
      ].join(','));
    }

    // Also check structure for leading elements work content
    if (tableArea.structure?.leadingElements?.workContent &&
        tableArea.structure.leadingElements.workContent.trim()) {

      const normCodes = tableArea.normCodes.join(',');

      rows.push([
        `"${normCodes}"`,
        this.escapeCsvField(tableArea.structure.leadingElements.workContent)
      ].join(','));
    }
  }

  // Generate 附注信息.csv - Notes information
  private generateFuZhuXinXi(): string {
    console.log('Generating 附注信息.csv...');
    const rows: string[] = [];

    // Header
    rows.push('编号,附注信息');

    // Process all chapters, sections, and table areas to find notes
    this.structuredData.chapters.forEach(chapter => {
      this.processChapterForFuZhuXinXi(chapter, rows);
    });

    return rows.join('\n');
  }

  private processChapterForFuZhuXinXi(chapter: Chapter, rows: string[]): void {
    // Process table areas directly in chapter
    chapter.tableAreas?.forEach(tableArea => {
      this.processTableAreaForFuZhuXinXi(tableArea, rows);
    });

    // Process sections
    chapter.sections.forEach(section => {
      this.processSectionForFuZhuXinXi(section, rows);
    });
  }

  private processSectionForFuZhuXinXi(section: Section, rows: string[]): void {
    // Process table areas directly in section
    section.tableAreas?.forEach(tableArea => {
      this.processTableAreaForFuZhuXinXi(tableArea, rows);
    });

    // Process subsections
    section.subSections.forEach(subSection => {
      this.processSubSectionForFuZhuXinXi(subSection, rows);
    });
  }

  private processSubSectionForFuZhuXinXi(subSection: SubSection, rows: string[]): void {
    // Process table areas in subsection
    subSection.tableAreas.forEach(tableArea => {
      this.processTableAreaForFuZhuXinXi(tableArea, rows);
    });

    // Process children subsections
    subSection.children.forEach(child => {
      this.processSubSectionForFuZhuXinXi(child, rows);
    });
  }

  private processTableAreaForFuZhuXinXi(tableArea: TableArea, rows: string[]): void {
    // Process notes from table area
    if (tableArea.notes && tableArea.notes.length > 0) {
      tableArea.notes.forEach(note => {
        if (note.trim()) {
          // Try to extract specific norm code from the note, or use all norms in the table
          const normCode = this.extractNormCodeFromNote(note) || tableArea.normCodes[0] || '';

          rows.push([
            this.escapeCsvField(normCode),
            this.escapeCsvField(note)
          ].join(','));
        }
      });
    }

    // Process notes from structure trailing elements
    if (tableArea.structure?.trailingElements?.notes &&
        tableArea.structure.trailingElements.notes.length > 0) {

      tableArea.structure.trailingElements.notes.forEach(note => {
        if (note.trim()) {
          const normCode = this.extractNormCodeFromNote(note) || tableArea.normCodes[0] || '';

          rows.push([
            this.escapeCsvField(normCode),
            this.escapeCsvField(note)
          ].join(','));
        }
      });
    }
  }

  private extractNormCodeFromNote(note: string): string | null {
    // Look for norm code pattern like "1B-1", "2B-16", etc. in the note
    const normCodeMatch = note.match(/\b\d+[A-Z]-\d+\b/);
    return normCodeMatch ? normCodeMatch[0] : null;
  }

  // Utility method to escape CSV fields
  private escapeCsvField(field: string): string {
    if (!field) return '';

    // If field contains comma, newline, or quote, wrap in quotes and escape internal quotes
    if (field.includes(',') || field.includes('\n') || field.includes('"')) {
      return `"${field.replace(/"/g, '""')}"`;
    }

    return field;
  }

  // Generate all CSV files
  public generateCsvFiles(outputDir: string): void {
    console.log(`Generating CSV files to: ${outputDir}`);

    // Ensure output directory exists
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Generate all CSV files
    const hanliangbiao = this.generateHanLiangBiao();
    fs.writeFileSync(path.join(outputDir, '含量表.csv'), hanliangbiao, 'utf8');

    const zimuxinxi = this.generateZiMuXinXi();
    fs.writeFileSync(path.join(outputDir, '子目信息.csv'), zimuxinxi, 'utf8');

    const gongzuoneirong = this.generateGongZuoNeiRong();
    fs.writeFileSync(path.join(outputDir, '工作内容.csv'), gongzuoneirong, 'utf8');

    const fuzhuxinxi = this.generateFuZhuXinXi();
    fs.writeFileSync(path.join(outputDir, '附注信息.csv'), fuzhuxinxi, 'utf8');

    console.log('All CSV files generated successfully!');
  }

  // Generate Excel files (3 files total)
  public async generateExcelFiles(outputDir: string): Promise<void> {
    console.log(`Generating Excel files to: ${outputDir}`);

    // Ensure output directory exists
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Generate 含量表.xlsx
    await this.generateHanLiangBiaoExcel(path.join(outputDir, '含量表.xlsx'));

    // Generate 子目信息.xlsx
    await this.generateZiMuXinXiExcel(path.join(outputDir, '子目信息.xlsx'));

    // Generate combined 工作内容和附注信息.xlsx with two sheets
    await this.generateCombinedWorkAndNotesExcel(path.join(outputDir, '工作内容和附注信息.xlsx'));

    console.log('All Excel files generated successfully!');
  }

  private async generateHanLiangBiaoExcel(filePath: string): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('含量表');

    // Add header
    const headers = ['编号', '名称', '规格', '单位', '单价', '含量', '主材标记', '材料号', '材料类别', '是否有明细', '', '', ''];
    worksheet.addRow(headers);

    // Add data
    const csvContent = this.generateHanLiangBiao();
    const rows = csvContent.split('\n');
    rows.slice(1).forEach(row => { // Skip header
      if (row.trim()) {
        const columns = this.parseCsvRow(row);
        worksheet.addRow(columns);
      }
    });

    await workbook.xlsx.writeFile(filePath);
  }

  private async generateZiMuXinXiExcel(filePath: string): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('子目信息');

    // Add header
    const headers = ['序号', '符号', '定额号', '子目名称', '单位', '基价', '人工', '材料', '机械', '图片名称'];
    worksheet.addRow(headers);

    // Add data
    const csvContent = this.generateZiMuXinXi();
    const rows = csvContent.split('\n');
    rows.slice(1).forEach(row => { // Skip header
      if (row.trim()) {
        const columns = this.parseCsvRow(row);
        worksheet.addRow(columns);
      }
    });

    await workbook.xlsx.writeFile(filePath);
  }

  private async generateCombinedWorkAndNotesExcel(filePath: string): Promise<void> {
    const workbook = new ExcelJS.Workbook();

    // Add 工作内容 sheet
    const workSheet = workbook.addWorksheet('工作内容');
    workSheet.addRow(['编号', '工作内容']);

    const workContent = this.generateGongZuoNeiRong();
    const workRows = workContent.split('\n');
    workRows.slice(1).forEach(row => {
      if (row.trim()) {
        const columns = this.parseCsvRow(row);
        workSheet.addRow(columns);
      }
    });

    // Add 附注信息 sheet
    const notesSheet = workbook.addWorksheet('附注信息');
    notesSheet.addRow(['编号', '附注信息']);

    const notesContent = this.generateFuZhuXinXi();
    const notesRows = notesContent.split('\n');
    notesRows.slice(1).forEach(row => {
      if (row.trim()) {
        const columns = this.parseCsvRow(row);
        notesSheet.addRow(columns);
      }
    });

    await workbook.xlsx.writeFile(filePath);
  }

  // Simple CSV parser for converting back to Excel
  private parseCsvRow(row: string): string[] {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < row.length; i++) {
      const char = row[i];

      if (char === '"' && !inQuotes) {
        inQuotes = true;
      } else if (char === '"' && inQuotes) {
        if (i + 1 < row.length && row[i + 1] === '"') {
          current += '"';
          i++; // Skip next quote
        } else {
          inQuotes = false;
        }
      } else if (char === ',' && !inQuotes) {
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }

    result.push(current);
    return result;
  }
}

// Main execution
async function main() {
  const inputPath = './output/structured-excel.json';
  const csvOutputDir = './output_final/csv';
  const excelOutputDir = './output_final/excel';

  try {
    const exporter = new FinalResultsExporter(inputPath);

    // Generate CSV files first for debugging
    exporter.generateCsvFiles(csvOutputDir);

    // Generate Excel files
    await exporter.generateExcelFiles(excelOutputDir);

    console.log('Final results export completed successfully!');
    console.log(`CSV files saved to: ${csvOutputDir}`);
    console.log(`Excel files saved to: ${excelOutputDir}`);

  } catch (error) {
    console.error('Error during final results export:', error);
    process.exit(1);
  }
}

// Run if called directly
if (require.main === module) {
  main();
}

export { FinalResultsExporter };
