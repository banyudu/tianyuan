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

  private isChapterTitle(value: string, cell?: CellData): boolean {
    const hasChapterPattern = /^第[一二三四五六七八九十\d]+章/.test(value);
    if (!cell || !hasChapterPattern) return hasChapterPattern;
    
    // Use font information to confirm chapter titles (SimHei font)
    return cell.font?.name === 'SimHei';
  }

  private isSectionTitle(value: string, cell?: CellData): boolean {
    // Only SimHei font cells can be section titles
    if (cell?.font?.name !== 'SimHei') {
      return false;
    }
    
    // True section titles like "第一节 减振装置安装"
    const hasSectionPattern = /^第[一二三四五六七八九十\d]+节/.test(value);
    return hasSectionPattern;
  }

  private isSubSectionTitle(value: string, cell?: CellData): boolean {
    // Only consider subsection titles that use SimHei font (structural headers)
    if (cell?.font?.name !== 'SimHei') {
      return false;
    }
    
    // Subsection titles with spaces like "一 、减振装置安装" use SimHei font  
    const hasSpacedSubSectionPattern = /^[一二三四五六七八九十]+\s+、/.test(value);
    if (hasSpacedSubSectionPattern) {
      return true;
    }
    
    // Numbered subsections like "(1)" - but these are rare and should also have SimHei
    return /^\(\d+\)/.test(value);
  }

  private isQuotaCode(value: string): boolean {
    return /^\d+[A-Z]-\d+$/.test(value);
  }

  private areInSameTable(quota1: {row: number; col: number}, quota2: {row: number; col: number}): boolean {
    // Check if two quota codes are in the same table by looking at the area between them
    const minRow = Math.min(quota1.row, quota2.row);
    const maxRow = Math.max(quota1.row, quota2.row);
    const minCol = Math.min(quota1.col, quota2.col);
    const maxCol = Math.max(quota1.col, quota2.col);
    
    // Much more restrictive - if they're too far apart, they're likely in different tables
    if (maxRow - minRow > 8 || maxCol - minCol > 8) {
      return false;
    }
    
    // Check if they're in the same row (horizontal table layout)
    if (quota1.row === quota2.row) {
      return true;
    }
    
    // For vertical proximity, check if there are connecting cells with borders
    let hasConnectingBorders = false;
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const cell = this.getCell(r, c);
        if (cell && (cell.borderStyles || cell.borders)) {
          hasConnectingBorders = true;
          break;
        }
      }
      if (hasConnectingBorders) break;
    }
    
    return hasConnectingBorders;
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
    // Match patterns like "第一节 减振装置安装"
    let match = value.match(/^(第[一二三四五六七八九十\d]+节)\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1],
        name: match[2].trim()
      };
    }
    
    // Match numbered patterns like "1.工地运输"
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
    // Match patterns like "一 、减振装置安装" (with spaces)
    let match = value.match(/^([一二三四五六七八九十]+)\s+、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + ' 、',
        name: match[2].trim()
      };
    }
    
    // Match patterns like "一、减振装置安装" (without spaces)
    match = value.match(/^([一二三四五六七八九十]+)、\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: match[1] + '、',
        name: match[2].trim()
      };
    }
    
    // Match patterns like "(1)单杆"
    match = value.match(/^\((\d+)\)\s*(.+?)(?:\s*·|$)/);
    if (match) {
      return {
        symbol: `(${match[1]})`,
        name: match[2].trim()
      };
    }
    
    return null;
  }

  private collectDescriptionText(startRow: number, endRow: number): string[] {
    const descriptions: string[] = [];
    
    for (let row = startRow; row <= endRow; row++) {
      const cellValue = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      
      // Only collect non-empty SimSun font cells (description text)
      if (cellValue && cell?.font?.name === 'SimSun') {
        // Skip if it looks like a structural element
        if (!this.isChapterTitle(cellValue, cell) && 
            !this.isSectionTitle(cellValue, cell) && 
            !this.isSubSectionTitle(cellValue, cell) &&
            !this.isQuotaCode(cellValue)) {
          descriptions.push(cellValue);
        }
      }
    }
    
    return descriptions;
  }

  private detectTableAreas(): TableArea[] {
    const tableAreas: TableArea[] = [];
    
    // Strategy 1: Find all quota codes first, then build tables around them
    const quotaCodeCells: Array<{row: number; col: number; code: string}> = [];
    
    for (const cell of this.data.cells) {
      if (cell.value && typeof cell.value === 'string' && this.isQuotaCode(cell.value)) {
        quotaCodeCells.push({
          row: cell.row,
          col: cell.col,
          code: cell.value
        });
      }
    }
    
    console.log(`Found ${quotaCodeCells.length} quota code cells`);
    
    // Group quota codes by proximity (same table)
    const processedQuotas = new Set<string>();
    
    for (const quotaCell of quotaCodeCells) {
      const key = `${quotaCell.row}-${quotaCell.col}`;
      if (processedQuotas.has(key)) continue;
      
      // Find the table boundaries around this quota code
      const tableQuotas: Array<{row: number; col: number; code: string}> = [];
      const visitedCells = new Set<string>();
      
      // Use BFS to find all connected quota codes in the same table structure
      const queue = [quotaCell];
      visitedCells.add(key);
      
      while (queue.length > 0) {
        const current = queue.shift()!;
        tableQuotas.push(current);
        
        // Look for nearby quota codes within reasonable distance (much smaller radius)
        const searchRadius = 8; // cells - reduced for more granular tables
        for (const otherQuota of quotaCodeCells) {
          const otherKey = `${otherQuota.row}-${otherQuota.col}`;
          if (visitedCells.has(otherKey)) continue;
          
          // Check if this quota is within the same table area
          const rowDistance = Math.abs(otherQuota.row - current.row);
          const colDistance = Math.abs(otherQuota.col - current.col);
          
          // More restrictive conditions for smaller tables
          if (rowDistance <= searchRadius && colDistance <= searchRadius) {
            // Additional check: ensure they're in the same bordered area and close enough
            if (this.areInSameTable(current, otherQuota) && rowDistance <= 5) {
              visitedCells.add(otherKey);
              queue.push(otherQuota);
            }
          }
        }
      }
      
      // Mark all found quotas as processed
      for (const tq of tableQuotas) {
        processedQuotas.add(`${tq.row}-${tq.col}`);
      }
      
      if (tableQuotas.length > 0) {
        // Calculate table boundaries
        const rows = tableQuotas.map(q => q.row);
        const cols = tableQuotas.map(q => q.col);
        
        const minRow = Math.min(...rows);
        const maxRow = Math.max(...rows);
        const minCol = Math.min(...cols);
        const maxCol = Math.max(...cols);
        
        // Expand boundaries to include table structure
        const startRow = Math.max(1, minRow - 2);
        const endRow = Math.min(this.data.metadata.totalRows, maxRow + 10);
        const startCol = Math.max(1, minCol - 2);
        const endCol = Math.min(this.data.metadata.totalCols, maxCol + 5);
        
        // Extract table information
        const quotaCodes = tableQuotas.map(q => q.code).sort();
        let unit = '';
        let workContent = '';
        const notes: string[] = [];
        
        // Look for work content in rows above the table
        for (let r = Math.max(1, startRow - 3); r < startRow + 3; r++) {
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
          if (workContent) break;
        }
        
        // Look for notes in rows below the table
        for (let r = maxRow; r <= Math.min(endRow + 5, this.data.metadata.totalRows); r++) {
          for (let c = startCol; c <= endCol; c++) {
            const value = this.getCellValue(r, c);
            if (this.isNote(value)) {
              notes.push(value);
            }
          }
        }
        
        const tableId = `table_${minRow}_${minCol}`;
        const table: TableArea = {
          id: tableId,
          range: { startRow, endRow, startCol, endCol },
          quotaCodes,
          unit,
          workContent: workContent || undefined, // Make optional
          notes,
          isContinuation: false
        };
        
        tableAreas.push(table);
        
        console.log(`  Found table at ${startRow}-${endRow}:${startCol}-${endCol} with ${quotaCodes.length} quotas: ${quotaCodes.slice(0, 5).join(', ')}${quotaCodes.length > 5 ? '...' : ''}`);
        if (workContent) console.log(`    Work content: ${workContent.substring(0, 50)}...`);
        if (notes.length > 0) console.log(`    Notes: ${notes.length} found`);
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

    // First pass: identify all structural headers
    const structuralHeaders: Array<{row: number, type: 'chapter' | 'section' | 'subsection', data: any}> = [];
    
    for (let row = 1; row <= this.data.metadata.totalRows; row++) {
      const cellValue = this.getCellValue(row, 1);
      if (!cellValue) continue;
      
      const cell = this.getCell(row, 1);

      if (this.isChapterTitle(cellValue, cell)) {
        const chapterInfo = this.parseChapterTitle(cellValue);
        if (chapterInfo) {
          structuralHeaders.push({row, type: 'chapter', data: chapterInfo});
        }
      } else if (this.isSectionTitle(cellValue, cell)) {
        const sectionInfo = this.parseSectionTitle(cellValue);
        if (sectionInfo) {
          structuralHeaders.push({row, type: 'section', data: sectionInfo});
        }
      } else if (this.isSubSectionTitle(cellValue, cell)) {
        const subSectionInfo = this.parseSubSectionTitle(cellValue);
        if (subSectionInfo) {
          structuralHeaders.push({row, type: 'subsection', data: subSectionInfo});
        }
      }
    }

    // Second pass: build hierarchy with descriptions
    for (let i = 0; i < structuralHeaders.length; i++) {
      const header = structuralHeaders[i];
      const nextHeader = structuralHeaders[i + 1];
      
      // Calculate description range (between current header and next header)
      const descriptionStart = header.row + 1;
      const descriptionEnd = nextHeader ? nextHeader.row - 1 : this.data.metadata.totalRows;
      
      if (header.type === 'chapter') {
        // Collect chapter description
        const description = this.collectDescriptionText(descriptionStart, Math.min(descriptionEnd, descriptionStart + 20));
        
        currentChapter = {
          id: `chapter_${header.data.number}`,
          name: header.data.name,
          number: header.data.number,
          description: description.length > 0 ? description : undefined,
          sections: [],
          tableAreas: []
        };
        chapters.push(currentChapter);
        currentSection = null;
        currentSubSection = null;
        this.processedRows.add(header.row);
        console.log(`Found chapter: ${header.data.name} (SimHei font confirmed)`);
        if (description.length > 0) {
          console.log(`  Chapter description: ${description.length} lines`);
        }
      } else if (header.type === 'section' && currentChapter) {
        // Collect section description
        const description = this.collectDescriptionText(descriptionStart, Math.min(descriptionEnd, descriptionStart + 15));
        
        currentSection = {
          id: `section_${currentChapter.id}_${header.data.symbol}`,
          name: header.data.name,
          number: header.data.symbol,
          description: description.length > 0 ? description : undefined,
          subSections: [],
          tableAreas: []
        };
        currentChapter.sections.push(currentSection);
        currentSubSection = null;
        this.processedRows.add(header.row);
        console.log(`Found section: ${header.data.name} (Font: SimHei)`);
        if (description.length > 0) {
          console.log(`  Section description: ${description.length} lines`);
        }
      } else if (header.type === 'subsection' && currentSection) {
        // Collect subsection description
        const description = this.collectDescriptionText(descriptionStart, Math.min(descriptionEnd, descriptionStart + 10));
        
        currentSubSection = {
          id: `subsection_${currentSection.id}_${header.data.symbol}`,
          name: header.data.name,
          level: 1,
          symbol: header.data.symbol,
          description: description.length > 0 ? description : undefined,
          tableAreas: [],
          children: []
        };
        currentSection.subSections.push(currentSubSection);
        this.processedRows.add(header.row);
        console.log(`Found subsection: ${header.data.name} (Font: SimHei)`);
        if (description.length > 0) {
          console.log(`  Subsection description: ${description.length} lines`);
        }
      }
    }

    // Assign table areas to appropriate hierarchy levels
    console.log(`Assigning ${tableAreas.length} table areas to hierarchy...`);
    
    for (const table of tableAreas) {
      const tableRow = table.range.startRow;
      let assigned = false;

      console.log(`Assigning table at row ${tableRow} with quotas: ${table.quotaCodes.join(', ')}`);

      // Find the most recent chapter/section/subsection above this table by walking backwards
      let bestChapter: Chapter | null = null;
      let bestSection: Section | null = null;
      let bestSubSection: SubSection | null = null;
      
      let lastChapterRow = -1;
      let lastSectionRow = -1;
      let lastSubSectionRow = -1;

      // Walk backwards from the table to find the most recent hierarchy headers
      for (let row = tableRow - 1; row >= 1; row--) {
        const cellValue = this.getCellValue(row, 1);
        const cell = this.getCell(row, 1);
        
        if (!cellValue) continue;

        // Check for subsection (most specific)
        if (lastSubSectionRow === -1 && this.isSubSectionTitle(cellValue, cell)) {
          lastSubSectionRow = row;
          // Find the corresponding subsection in our hierarchy
          for (const chapter of chapters) {
            for (const section of chapter.sections) {
              for (const subSection of section.subSections) {
                if (cellValue.includes(subSection.name) || subSection.name.includes(cellValue.substring(0, 10))) {
                  bestSubSection = subSection;
                  bestSection = section;
                  bestChapter = chapter;
                  break;
                }
              }
              if (bestSubSection) break;
            }
            if (bestSubSection) break;
          }
        }
        
        // Check for section (medium specificity)
        if (lastSectionRow === -1 && this.isSectionTitle(cellValue, cell)) {
          lastSectionRow = row;
          // Only update if we haven't found a subsection
          if (!bestSubSection) {
            for (const chapter of chapters) {
              for (const section of chapter.sections) {
                if (cellValue.includes(section.name) || section.name.includes(cellValue.substring(0, 10))) {
                  bestSection = section;
                  bestChapter = chapter;
                  break;
                }
              }
              if (bestSection) break;
            }
          }
        }
        
        // Check for chapter (least specific)
        if (lastChapterRow === -1 && this.isChapterTitle(cellValue, cell)) {
          lastChapterRow = row;
          // Only update if we haven't found a section or subsection
          if (!bestSection && !bestSubSection) {
            for (const chapter of chapters) {
              if (cellValue.includes(chapter.name) || chapter.name.includes(cellValue.substring(0, 10))) {
                bestChapter = chapter;
                break;
              }
            }
          }
        }

        // If we've found all levels or gone too far back, stop searching
        if ((bestSubSection || bestSection || bestChapter) && row < tableRow - 50) {
          break;
        }
      }

      // Log what we found
      console.log(`  Found headers above table at row ${tableRow}:`);
      console.log(`    Chapter: ${bestChapter?.name || 'None'} (row ${lastChapterRow})`);
      console.log(`    Section: ${bestSection?.name || 'None'} (row ${lastSectionRow})`);
      console.log(`    SubSection: ${bestSubSection?.name || 'None'} (row ${lastSubSectionRow})`);

      // Assign to the most specific level found (prefer leaf nodes - subSections)
      if (bestSubSection) {
        bestSubSection.tableAreas.push(table);
        console.log(`  -> Assigned to subsection: ${bestSubSection.name}`);
        assigned = true;
      } else if (bestSection) {
        // If no subsection, assign to section (but prefer creating a default subsection)
        if (bestSection.subSections.length === 0) {
          // Create a default subsection for tables without explicit subsections
          const defaultSubSection: SubSection = {
            id: `subsection_${bestSection.id}_default`,
            name: 'Tables',
            level: 1,
            symbol: '',
            tableAreas: [table],
            children: []
          };
          bestSection.subSections.push(defaultSubSection);
          console.log(`  -> Created default subsection and assigned table to section: ${bestSection.name}`);
        } else {
          // Assign to the last subsection in this section
          const lastSubSection = bestSection.subSections[bestSection.subSections.length - 1];
          lastSubSection.tableAreas.push(table);
          console.log(`  -> Assigned to last subsection: ${lastSubSection.name} in section: ${bestSection.name}`);
        }
        assigned = true;
      } else if (bestChapter) {
        // If no section, assign to chapter (but prefer creating a default section/subsection)
        if (bestChapter.sections.length === 0) {
          // Create a default section and subsection
          const defaultSection: Section = {
            id: `section_${bestChapter.id}_default`,
            name: 'Default Section',
            number: '',
            subSections: [],
            tableAreas: []
          };
          
          const defaultSubSection: SubSection = {
            id: `subsection_${defaultSection.id}_default`,
            name: 'Tables',
            level: 1,
            symbol: '',
            tableAreas: [table],
            children: []
          };
          
          defaultSection.subSections.push(defaultSubSection);
          bestChapter.sections.push(defaultSection);
          console.log(`  -> Created default section/subsection and assigned to chapter: ${bestChapter.name}`);
        } else {
          // Assign to the last section's last subsection
          const lastSection = bestChapter.sections[bestChapter.sections.length - 1];
          if (lastSection.subSections.length === 0) {
            const defaultSubSection: SubSection = {
              id: `subsection_${lastSection.id}_default`,
              name: 'Tables',
              level: 1,
              symbol: '',
              tableAreas: [table],
              children: []
            };
            lastSection.subSections.push(defaultSubSection);
            console.log(`  -> Created default subsection in last section: ${lastSection.name}`);
          } else {
            const lastSubSection = lastSection.subSections[lastSection.subSections.length - 1];
            lastSubSection.tableAreas.push(table);
            console.log(`  -> Assigned to last subsection: ${lastSubSection.name}`);
          }
        }
        assigned = true;
      }

      if (!assigned) {
        console.log(`  -> WARNING: Could not assign table at row ${tableRow}`);
        // Create a fallback structure
        if (chapters.length > 0) {
          const firstChapter = chapters[0];
          if (firstChapter.sections.length === 0) {
            const defaultSection: Section = {
              id: `section_${firstChapter.id}_fallback`,
              name: 'Fallback Section',
              number: '',
              subSections: [],
              tableAreas: []
            };
            firstChapter.sections.push(defaultSection);
          }
          
          const lastSection = firstChapter.sections[firstChapter.sections.length - 1];
          if (lastSection.subSections.length === 0) {
            const defaultSubSection: SubSection = {
              id: `subsection_${lastSection.id}_fallback`,
              name: 'Fallback Tables',
              level: 1,
              symbol: '',
              tableAreas: [],
              children: []
            };
            lastSection.subSections.push(defaultSubSection);
          }
          
          const lastSubSection = lastSection.subSections[lastSection.subSections.length - 1];
          lastSubSection.tableAreas.push(table);
          console.log(`  -> Fallback: Assigned to subsection in first chapter`);
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
      const cell = this.getCell(row, 1);
      if (this.isChapterTitle(value, cell)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextSectionRow(currentRow: number, chapter: Chapter): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      if (this.isSectionTitle(value, cell) || this.isChapterTitle(value, cell)) {
        return row;
      }
    }
    return this.data.metadata.totalRows;
  }

  private findNextSubSectionRow(currentRow: number, section: Section): number {
    for (let row = currentRow + 1; row <= this.data.metadata.totalRows; row++) {
      const value = this.getCellValue(row, 1);
      const cell = this.getCell(row, 1);
      if (this.isSubSectionTitle(value, cell) || this.isSectionTitle(value, cell) || this.isChapterTitle(value, cell)) {
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