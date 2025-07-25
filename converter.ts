import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

// Data models based on analysis
interface SubitemInfo {
    code: string;
    name: string;
    unit: string;
    workContent?: string;
    specification?: string;
}

interface WorkContent {
    code: string;
    name: string;
    workContent: string;
    unit: string;
    notes?: string;
}

interface MaterialContent {
    subitemCode: string;
    materialName: string;
    unit: string;
    quantity: number;
    category: string; // 人工, 材料, 机械
}

class ExcelConverter {
    private workbook: ExcelJS.Workbook;
    private worksheet: ExcelJS.Worksheet | null = null;

    constructor() {
        this.workbook = new ExcelJS.Workbook();
    }

    private getCellValue(cell: ExcelJS.Cell): string {
        if (!cell.value) return '';

        // Handle rich text
        if (typeof cell.value === 'object' && 'richText' in cell.value) {
            return cell.value.richText.map((rt: any) => rt.text).join('');
        }

        // Handle formulas
        if (typeof cell.value === 'object' && 'formula' in cell.value) {
            return cell.value.result?.toString() || '';
        }

        // Handle shared strings
        if (typeof cell.value === 'object' && 'sharedString' in cell.value) {
            return (cell.value as any).sharedString.toString();
        }

        return cell.value.toString();
    }

    async loadInputFile(filePath: string): Promise<void> {
        console.log(`Loading input file: ${filePath}`);
        await this.workbook.xlsx.readFile(filePath);
        this.worksheet = this.workbook.getWorksheet(1) || null;

        if (!this.worksheet) {
            throw new Error('No worksheet found in the input file');
        }

        console.log(`Loaded worksheet with ${this.worksheet.rowCount} rows and ${this.worksheet.columnCount} columns`);
    }

    private findSectionBoundaries(): { subitemStart: number, workContentStart: number, materialStart: number } {
        if (!this.worksheet) throw new Error('Worksheet not loaded');

        let subitemStart = 0;
        let workContentStart = 0;
        let materialStart = 0;

        // Find section boundaries based on analysis
        for (let row = 1; row <= this.worksheet.rowCount; row++) {
            const rowText: string[] = [];
            for (let col = 1; col <= Math.min(15, this.worksheet.columnCount); col++) {
                const cell = this.worksheet.getCell(row, col);
                const value = this.getCellValue(cell);
                rowText.push(value);
            }
            const rowContent = rowText.join(' ');

            // Look for subitem section
            if (!subitemStart && rowContent.includes('子目编号')) {
                subitemStart = row;
                console.log(`Found subitem section at row ${row}`);
            }

            // Look for work content section
            if (!workContentStart && rowContent.includes('工作内容：') && row > 500) {
                workContentStart = row;
                console.log(`Found work content section at row ${row}`);
            }

            // Look for material section (starts early in file)
            if (!materialStart && (rowContent.includes('材料') && rowContent.includes('名') && row > 50)) {
                materialStart = row;
                console.log(`Found material section at row ${row}`);
            }
        }

        return { subitemStart, workContentStart, materialStart };
    }

    private extractSubitemInfo(startRow: number): SubitemInfo[] {
        if (!this.worksheet) return [];

        const subitems: SubitemInfo[] = [];
        let currentSubitem: Partial<SubitemInfo> = {};

        console.log(`Extracting subitem info starting from row ${startRow}`);

        for (let row = startRow; row <= Math.min(startRow + 200, this.worksheet.rowCount); row++) {
            const rowValues: string[] = [];
            for (let col = 1; col <= Math.min(10, this.worksheet.columnCount); col++) {
                const cell = this.worksheet.getCell(row, col);
                const value = this.getCellValue(cell);
                rowValues.push(value);
            }

            // Look for subitem codes (pattern like 1B-1, 2A-1, etc.)
            const codePattern = /([0-9]+[A-Z]+-[0-9]+)/;
            const codeMatch = rowValues.join(' ').match(codePattern);

            if (codeMatch) {
                // Save previous subitem if exists
                if (currentSubitem.code) {
                    subitems.push(currentSubitem as SubitemInfo);
                }

                currentSubitem = {
                    code: codeMatch[1],
                    name: '',
                    unit: '',
                    workContent: ''
                };

                // Look for name and unit in the same row or nearby
                for (const value of rowValues) {
                    if (value && value.length > 5 && /[\u4e00-\u9fff]/.test(value) && !currentSubitem.name) {
                        currentSubitem.name = value;
                    }
                    if (/^(个|台|套|m|kg|只|根|块|张|副|m²|m³)$/.test(value)) {
                        currentSubitem.unit = value;
                    }
                }
            } else if (currentSubitem.code) {
                // Continue building current subitem
                for (const value of rowValues) {
                    if (value && value.length > 5 && /[\u4e00-\u9fff]/.test(value) && !currentSubitem.name) {
                        currentSubitem.name = value;
                    }
                    if (/^(个|台|套|m|kg|只|根|块|张|副|m²|m³)$/.test(value) && !currentSubitem.unit) {
                        currentSubitem.unit = value;
                    }
                    if (value.includes('工作内容') && !currentSubitem.workContent) {
                        currentSubitem.workContent = value;
                    }
                }
            }
        }

        // Save last subitem
        if (currentSubitem.code) {
            subitems.push(currentSubitem as SubitemInfo);
        }

        console.log(`Extracted ${subitems.length} subitems`);
        return subitems;
    }

    private extractWorkContent(startRow: number): WorkContent[] {
        if (!this.worksheet) return [];

        const workContents: WorkContent[] = [];

        console.log(`Extracting work content starting from row ${startRow}`);

        for (let row = startRow; row <= this.worksheet.rowCount; row++) {
            const rowValues: string[] = [];
            for (let col = 1; col <= Math.min(10, this.worksheet.columnCount); col++) {
                const cell = this.worksheet.getCell(row, col);
                const value = this.getCellValue(cell);
                rowValues.push(value);
            }

            const rowText = rowValues.join(' ');

            // Look for work content patterns
            if (rowText.includes('工作内容：') && rowText.includes('单位')) {
                const workContent: Partial<WorkContent> = {};

                // Extract work content description
                const contentMatch = rowText.match(/工作内容：([^。]+)/);
                if (contentMatch) {
                    workContent.workContent = contentMatch[1];
                }

                // Extract unit
                const unitMatch = rowText.match(/单位\s*：\s*([^\s]+)/);
                if (unitMatch) {
                    workContent.unit = unitMatch[1];
                }

                // Generate a code based on position
                workContent.code = `WC-${workContents.length + 1}`;
                workContent.name = workContent.workContent?.slice(0, 20) || '';

                if (workContent.workContent && workContent.unit) {
                    workContents.push(workContent as WorkContent);
                }
            }
        }

        console.log(`Extracted ${workContents.length} work content items`);
        return workContents;
    }

    private extractMaterialContent(startRow: number): MaterialContent[] {
        if (!this.worksheet) return [];

        const materials: MaterialContent[] = [];
        let currentSubitemCode = '';

        console.log(`Extracting material content starting from row ${startRow}`);

        for (let row = startRow; row <= this.worksheet.rowCount; row++) {
            const rowValues: string[] = [];
            for (let col = 1; col <= Math.min(15, this.worksheet.columnCount); col++) {
                const cell = this.worksheet.getCell(row, col);
                const value = this.getCellValue(cell);
                rowValues.push(value);
            }

            const rowText = rowValues.join(' ');

            // Look for subitem codes to track context
            const codePattern = /([0-9]+[A-Z]+-[0-9]+)/;
            const codeMatch = rowText.match(codePattern);
            if (codeMatch) {
                currentSubitemCode = codeMatch[1];
            }

            // Look for material entries
            if (rowText.includes('材料') && currentSubitemCode) {
                // Extract material names and quantities
                for (let i = 0; i < rowValues.length; i++) {
                    const value = rowValues[i];
                    if (value && value.length > 3 && /[\u4e00-\u9fff]/.test(value) &&
                        !value.includes('材料') && !value.includes('名称')) {

                        // Look for quantity in nearby cells
                        let quantity = 0;
                        for (let j = Math.max(0, i-2); j <= Math.min(rowValues.length-1, i+2); j++) {
                            const numValue = parseFloat(rowValues[j]);
                            if (!isNaN(numValue) && numValue > 0) {
                                quantity = numValue;
                                break;
                            }
                        }

                        materials.push({
                            subitemCode: currentSubitemCode,
                            materialName: value,
                            unit: '个', // Default unit
                            quantity: quantity || 1,
                            category: '材料'
                        });
                    }
                }
            }
        }

        console.log(`Extracted ${materials.length} material items`);
        return materials;
    }

    private async createOutputWorkbook(data: SubitemInfo[] | WorkContent[] | MaterialContent[], fileName: string): Promise<void> {
        const outputWorkbook = new ExcelJS.Workbook();
        const worksheet = outputWorkbook.addWorksheet('Sheet1');

        if (fileName.includes('子目信息')) {
            // Subitem info table
            const subitems = data as SubitemInfo[];
            worksheet.addRow(['子目编号', '子目名称', '计量单位', '工作内容']);

            subitems.forEach(item => {
                worksheet.addRow([item.code, item.name, item.unit, item.workContent || '']);
            });

        } else if (fileName.includes('工作内容')) {
            // Work content table
            const workContents = data as WorkContent[];
            worksheet.addRow(['编号', '名称', '工作内容', '计量单位', '备注']);

            workContents.forEach(item => {
                worksheet.addRow([item.code, item.name, item.workContent, item.unit, item.notes || '']);
            });

        } else if (fileName.includes('含量表')) {
            // Material content table
            const materials = data as MaterialContent[];
            worksheet.addRow(['子目编号', '材料名称', '计量单位', '消耗量', '类别']);

            materials.forEach(item => {
                worksheet.addRow([item.subitemCode, item.materialName, item.unit, item.quantity, item.category]);
            });
        }

        // Save as .xls format
        await outputWorkbook.xlsx.writeFile(fileName);
        console.log(`Created output file: ${fileName}`);
    }

    async convert(inputFile: string, outputDir: string): Promise<void> {
        try {
            // Load input file
            await this.loadInputFile(inputFile);

            // Find section boundaries
            const sections = this.findSectionBoundaries();

            // Extract data from each section
            const subitemInfo = this.extractSubitemInfo(sections.subitemStart);
            const workContent = this.extractWorkContent(sections.workContentStart);
            const materialContent = this.extractMaterialContent(sections.materialStart);

            // Ensure output directory exists
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }

            // Create output files
            await this.createOutputWorkbook(subitemInfo, path.join(outputDir, '子目信息.xls'));
            await this.createOutputWorkbook(workContent, path.join(outputDir, '工作内容、附注信息表.xls'));
            await this.createOutputWorkbook(materialContent, path.join(outputDir, '含量表.xls'));

            console.log('\nConversion completed successfully!');
            console.log(`Output files created in: ${outputDir}`);

        } catch (error) {
            console.error('Conversion failed:', error);
            throw error;
        }
    }
}

// Main execution
async function main() {
    const inputFile = 'data/建设工程消耗量标准及计算规则（安装工程） 补充子目.xlsx';
    const outputDir = 'output';

    const converter = new ExcelConverter();
    await converter.convert(inputFile, outputDir);
}

// Run if this file is executed directly
if (require.main === module) {
    main().catch(console.error);
}
