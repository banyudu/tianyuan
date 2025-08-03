import * as fs from 'fs';
import * as path from 'path';
import { CellData, ParsedExcelData } from './src/types';
import { 
  StructuredDocument, 
  Chapter, 
  Section, 
  SubSection, 
  TableArea, 
  TableRange, 
  CellInfo, 
  BorderInfo 
} from './src/structured-types';

class StructuredExcelParser {
  private data: ParsedExcelData;
  private cellMap: Map<string, CellData>;
  private processedRows: Set<number>;

  constructor(jsonData: ParsedExcelData) {
    this.data = jsonData;
    this.cellMap = new Map();
    this.processedRows = new Set();
    
    // Create cell lookup map
    for (const cell of this.data.cells) {
      this.cellMap.set(`${cell.row}-${cell.col}`, cell);
      this.cellMap.set(cell.address, cell);
    }
  }

  private getCell(row: number, col: number): CellData | undefined {
    return this.cellMap.get(`${row}-${col}`);
  }

  private getCellValue(row: number, col: number): string {
    const cell = this.getCell(row, col);
    if (!cell || !cell.value) return '';
    return String(cell.value).trim();
  }

  private getBorderInfo(cell: CellData): BorderInfo {
    const borders = cell.borderStyles || {};
    return {
      hasTop: !!(borders.top?.style),
      hasBottom: !!(borders.bottom?.style),
      hasLeft: !!(borders.left?.style),
      hasRight: !!(borders.right?.style),
      topStyle: borders.top?.style,
      bottomStyle: borders.bottom?.style,
      leftStyle: borders.left?.style,
      rightStyle: borders.right?.style
    };
  }

  private isChapterTitle(value: string): boolean {
    return /^第[一二三四五六七八九十\d]+章/.test(value);
  }

  private isSectionTitle(value: string): boolean {
    return /^[一二三四五六七八九十]+、/.test(value) || /^\d+\./.test(value);
  }

  private isSubSectionTitle(value: string): boolean {
    return /^\(\d+\)/.test(value) || /^[一二三四五六七八九十]+\./.test(value);
  }

  private isQuotaCode(value: string): boolean {
    return /^\d+[A-Z]-\d+$/.test(value);
  }

  private isWorkContent(value: string): boolean {
    return value.includes('工作') && value.includes('内容') && value.includes('：');
  }

  private isNote(value: string): boolean {
    return value.startsWith('注') && (value.includes(':') || value.includes('：'));
  }

  private isContinuationTable(value: string): boolean {
    return value.includes('续表') || value.includes('（续）') || value.includes('(续)');
  }

  private parseChapterTitle(value: string): { number: string; name: string } | null {
    const match = value.match(/^第([一二三四五六七八九十\d]+)章\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        number: match[1],
        name: match[2].trim()
      };
    }
    return null;
  }

  private parseSectionTitle(value: string): { symbol: string; name: string } | null {
    // Match patterns like "一、机械设备安装工程" or "1.工地运输"
    let match = value.match(/^([一二三四五六七八九十]+)、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '、',
        name: match[2].trim()
      };
    }
    
    match = value.match(/^(\d+)\.\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '.',
        name: match[2].trim()
      };
    }
    
    return null;
  }

  private parseSubSectionTitle(value: string): { symbol: string; name: string } | null {
    // Match patterns like "(1)单杆" or "1.工地运输"
    let match = value.match(/^\((\d+)\)\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: `(${match[1]})`,
        name: match[2].trim()
      };
    }
    
    return null;
  }

  private detectTableAreas(): TableArea[] {
    const tableAreas: TableArea[] = [];
    const processedCells = new Set<string>();

    // Find cells with medium or thick borders to detect table regions
    for (const cell of this.data.cells) {
      const key = `${cell.row}-${cell.col}`;
      if (processedCells.has(key)) continue;

      // Check if cell has border styles (if borderStyles exists in data)
      let hasMediumBorder = false;
      if (cell.borderStyles) {
        const borderInfo = this.getBorderInfo(cell);
        hasMediumBorder = borderInfo.topStyle === 'medium' || 
                         borderInfo.bottomStyle === 'medium' || 
                         borderInfo.leftStyle === 'medium' || 
                         borderInfo.rightStyle === 'medium';
      } else if (cell.borders) {
        // Fallback to old border format
        hasMediumBorder = cell.borders.top || cell.borders.bottom || 
                         cell.borders.left || cell.borders.right;
      }

      if (!hasMediumBorder) continue;

      // Find table boundaries by expanding from this cell
      const startRow = cell.row;
      let endRow = startRow;
      let startCol = cell.col;
      let endCol = startCol;

      const maxSearchRows = 25;
      const maxSearchCols = 32;

      // Expand to find the full table area
      for (let r = startRow; r <= startRow + maxSearchRows && r <= this.data.metadata.totalRows; r++) {
        for (let c = startCol; c <= startCol + maxSearchCols && c <= this.data.metadata.totalCols; c++) {
          const checkCell = this.getCell(r, c);
          if (checkCell) {
            const checkBorderInfo = this.getBorderInfo(checkCell);
            const hasAnyMediumBorder = checkBorderInfo.topStyle === 'medium' || 
                                     checkBorderInfo.bottomStyle === 'medium' || 
                                     checkBorderInfo.leftStyle === 'medium' || 
                                     checkBorderInfo.rightStyle === 'medium';
            
            if (hasAnyMediumBorder) {
              endRow = Math.max(endRow, r);
              endCol = Math.max(endCol, c);
            }
          }
        }
      }

      // Extract quota codes, work content, and notes for this table
      const quotaCodes: string[] = [];
      let unit = '';
      let workContent = '';
      const notes: string[] = [];

      // Look for quota codes in the table header area (first few rows)
      for (let r = startRow; r <= Math.min(startRow + 3, endRow); r++) {
        for (let c = startCol; c <= endCol; c++) {
          const value = this.getCellValue(r, c);
          if (this.isQuotaCode(value)) {
            quotaCodes.push(value);
          }
        }
      }

      // Look for work content in rows above the table
      for (let r = Math.max(1, startRow - 3); r < startRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const value = this.getCellValue(r, c);
          if (this.isWorkContent(value)) {
            workContent = value;
            break;
          }
          if (value && value.includes('单位') && value.includes('：')) {
            unit = value;
          }
        }
      }

      // Look for notes in rows below the table
      for (let r = endRow + 1; r <= Math.min(endRow + 5, this.data.metadata.totalRows); r++) {
        for (let c = startCol; c <= endCol; c++) {
          const value = this.getCellValue(r, c);
          if (this.isNote(value)) {
            notes.push(value);
          }
        }
      }

      if (quotaCodes.length > 0) {
        const tableId = `table_${startRow}_${startCol}`;
        const table: TableArea = {
          id: tableId,
          range: { startRow, endRow, startCol, endCol },
          quotaCodes,
          unit,
          workContent,
          notes,
          isContinuation: false
        };
        
        tableAreas.push(table);
        
        console.log(`  Found table at ${startRow}-${endRow}:${startCol}-${endCol} with ${quotaCodes.length} quotas: ${quotaCodes.join(', ')}`);
        if (workContent) console.log(`    Work content: ${workContent.substring(0, 50)}...`);
        if (notes.length > 0) console.log(`    Notes: ${notes.length} found`);

        // Mark cells as processed
        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            processedCells.add(`${r}-${c}`);
          }
        }
      } else {
        // Debug: table found but no quota codes
        if (endRow - startRow > 3 && endCol - startCol > 3) {
          console.log(`  Found table structure at ${startRow}-${endRow}:${startCol}-${endCol} but no quota codes detected`);
        }
      }
    }

    // Process continuation tables
    this.processContinuationTables(tableAreas);

    return tableAreas;
  }

  private processContinuationTables(tableAreas: TableArea[]): void {
    for (let i = 0; i < tableAreas.length; i++) {
      const table = tableAreas[i];
      
      // Check if there's a "续表" indicator near this table
      for (let r = table.range.startRow - 2; r <= table.range.startRow + 2; r++) {
        const value = this.getCellValue(r, 1);
        if (this.isContinuationTable(value)) {
          table.isContinuation = true;
          
          // Find the previous table with the same quota codes
          for (let j = i - 1; j >= 0; j--) {
            const prevTable = tableAreas[j];
            if (prevTable.quotaCodes.some(code => table.quotaCodes.includes(code))) {
              table.continuationOf = prevTable.id;
              break;
            }
          }
          break;
        }
      }
    }
  }

  private buildHierarchicalStructure(tableAreas: TableArea[]): Chapter[] {
    const chapters: Chapter[] = [];
    let currentChapter: Chapter | null = null;
    let currentSection: Section | null = null;
    let currentSubSection: SubSection | null = null;

    // Scan through all rows to build hierarchy
    for (let row = 1; row <= this.data.metadata.totalRows; row++) {
      if (this.processedRows.has(row)) continue;

      const cellValue = this.getCellValue(row, 1);
      if (!cellValue) continue;

      // Check for chapter
      if (this.isChapterTitle(cellValue)) {
        const chapterInfo = this.parseChapterTitle(cellValue);
        if (chapterInfo) {
          currentChapter = {
            id: `chapter_${chapterInfo.number}`,
            name: chapterInfo.name,
            number: chapterInfo.number,
            sections: [],
            tableAreas: []
          };
          chapters.push(currentChapter);
          currentSection = null;
          currentSubSection = null;
          this.processedRows.add(row);
        }
      }
      // Check for section
      else if (this.isSectionTitle(cellValue)) {
        const sectionInfo = this.parseSectionTitle(cellValue);
        if (sectionInfo && currentChapter) {
          currentSection = {
            id: `section_${currentChapter.id}_${sectionInfo.symbol}`,
            name: sectionInfo.name,
            number: sectionInfo.symbol,
            subSections: [],
            tableAreas: []
          };
          currentChapter.sections.push(currentSection);
          currentSubSection = null;
          this.processedRows.add(row);
        }
      }
      // Check for sub-section
      else if (this.isSubSectionTitle(cellValue)) {
        const subSectionInfo = this.parseSubSectionTitle(cellValue);
        if (subSectionInfo && currentSection) {
          currentSubSection = {
            id: `subsection_${currentSection.id}_${subSectionInfo.symbol}`,
            name: subSectionInfo.name,
            level: 1,
            symbol: subSectionInfo.symbol,
            tableAreas: [],
            children: []
          };
          currentSection.subSections.push(currentSubSection);
          this.processedRows.add(row);
        }
      }
    }

    // Assign table areas to appropriate hierarchy levels
    console.log(`Assigning ${tableAreas.length} table areas to hierarchy...`);
    
    for (const table of tableAreas) {
      const tableRow = table.range.startRow;
      let assigned = false;

      console.log(`Assigning table at row ${tableRow} with quotas: ${table.quotaCodes.join(', ')}`);

      // Find the appropriate hierarchy level for this table
      let bestChapter: Chapter | null = null;
      let bestSection: Section | null = null;
      let bestSubSection: SubSection | null = null;

      // Find the closest chapter
      for (const chapter of chapters) {
        const chapterStartRow = this.findRowForText(chapter.name);
        const nextChapterStartRow = this.findNextChapterRow(chapterStartRow);
        
        if (tableRow > chapterStartRow && tableRow < nextChapterStartRow) {
          bestChapter = chapter;
          
          // Find the closest section within this chapter
          for (const section of chapter.sections) {
            const sectionStartRow = this.findRowForText(section.name);
            const nextSectionStartRow = this.findNextSectionRow(sectionStartRow, chapter);
            
            if (tableRow > sectionStartRow && tableRow < nextSectionStartRow) {
              bestSection = section;
              
              // Find the closest subsection within this section
              for (const subSection of section.subSections) {
                const subSectionStartRow = this.findRowForText(subSection.name);
                const nextSubSectionStartRow = this.findNextSubSectionRow(subSectionStartRow, section);
                
                if (tableRow > subSectionStartRow && tableRow < nextSubSectionStartRow) {
                  bestSubSection = subSection;
                  break;
                }
              }
              break;
            }
          }
          break;
        }
      }

      // Assign to the most specific level found
      if (bestSubSection) {
        bestSubSection.tableAreas.push(table);
        console.log(`  -> Assigned to subsection: ${bestSubSection.name}`);
        assigned = true;
      } else if (bestSection) {
        bestSection.tableAreas.push(table);
        console.log(`  -> Assigned to section: ${bestSection.name}`);
        assigned = true;
      } else if (bestChapter) {
        bestChapter.tableAreas.push(table);
        console.log(`  -> Assigned to chapter: ${bestChapter.name}`);
        assigned = true;
      }

      if (!assigned) {
        console.log(`  -> WARNING: Could not assign table at row ${tableRow}`);
        // Assign to first chapter as fallback
        if (chapters.length > 0) {
          chapters[0].tableAreas.push(table);
          console.log(`  -> Fallback: Assigned to first chapter: ${chapters[0].name}`);
        }
      }
    }

    return chapters;
  }

  private findRowForText(text: string): number {
    for (let row = 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      if (value.includes(text)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextChapterRow(currentRow: number): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      if (this.isChapterTitle(value)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextSectionRow(currentRow: number, chapter: Chapter): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      if (this.isSectionTitle(value) || this.isChapterTitle(value)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextSubSectionRow(currentRow: number, section: Section): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      if (this.isSubSectionTitle(value) || this.isSectionTitle(value) || this.isChapterTitle(value)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  public parseStructure(): StructuredDocument {
    console.log('Starting structured parsing...');
    
    // Step 1: Detect all table areas
    console.log('Detecting table areas...');
    const tableAreas = this.detectTableAreas();
    console.log(`Found ${tableAreas.length} table areas`);

    // Step 2: Build hierarchical structure
    console.log('Building hierarchical structure...');
    const chapters = this.buildHierarchicalStructure(tableAreas);
    console.log(`Created ${chapters.length} chapters`);

    // Step 3: Create structured document
    const structuredDoc: StructuredDocument = {
      metadata: {
        ...this.data.metadata,
        structuredAt: new Date().toISOString()
      },
      chapters
    };

    console.log('Structured parsing completed');
    return structuredDoc;
  }
}

// Main execution
async function main() {
  const inputPath = './output/parsed-excel.json';
  const outputPath = './output/structured-excel.json';
  
  try {
    console.log(`Loading JSON data from: ${inputPath}`);
    const jsonData = JSON.parse(fs.readFileSync(inputPath, 'utf8')) as ParsedExcelData;
    
    const parser = new StructuredExcelParser(jsonData);
    const structuredDoc = parser.parseStructure();
    
    // Ensure output directory exists
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
    
    // Write structured document
    fs.writeFileSync(outputPath, JSON.stringify(structuredDoc, null, 2), 'utf8');
    console.log(`Structured document saved to: ${outputPath}`);
    
    // Print summary
    console.log('\n=== PARSING SUMMARY ===');
    console.log(`Total chapters: ${structuredDoc.chapters.length}`);
    
    for (const chapter of structuredDoc.chapters) {
      console.log(`\n📖 Chapter ${chapter.number}: ${chapter.name}`);
      console.log(`   Sections: ${chapter.sections.length}, Tables: ${chapter.tableAreas.length}`);
      
      for (const section of chapter.sections) {
        console.log(`   📑 ${section.number} ${section.name}`);
        console.log(`      SubSections: ${section.subSections.length}, Tables: ${section.tableAreas.length}`);
        
        for (const subSection of section.subSections) {
          console.log(`      📄 ${subSection.symbol} ${subSection.name} (Tables: ${subSection.tableAreas.length})`);
        }
      }
    }
    
  } catch (error) {
    console.error('Error during structured parsing:', error);
    process.exit(1);
  }
}

// Run if called directly
if (require.main === module) {
  main();
}

export { StructuredExcelParser };