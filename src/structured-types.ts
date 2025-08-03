export interface TableRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface TableArea {
  id: string;
  range: TableRange;
  quotaCodes: string[];
  unit?: string;
  workContent?: string;  // Optional - not every table has work content
  notes: string[];       // Optional - not every table has notes
  isContinuation?: boolean;
  continuationOf?: string;
}

export interface SubSection {
  id: string;
  name: string;
  level: number; // 1, 2, 3, 4 for different hierarchy levels
  symbol: string; // "一、", "1.", "(1)", etc.
  tableAreas: TableArea[];
  children: SubSection[];
}

export interface Section {
  id: string;
  name: string;
  number: string;
  subSections: SubSection[];
  tableAreas: TableArea[];
}

export interface Chapter {
  id: string;
  name: string;
  number: string;
  sections: Section[];
  tableAreas: TableArea[];
}

export interface StructuredDocument {
  metadata: {
    filename: string;
    sheetName: string;
    totalRows: number;
    totalCols: number;
    parsedAt: string;
    structuredAt: string;
  };
  chapters: Chapter[];
}

export interface BorderInfo {
  hasTop: boolean;
  hasBottom: boolean;
  hasLeft: boolean;
  hasRight: boolean;
  topStyle?: string;
  bottomStyle?: string;
  leftStyle?: string;
  rightStyle?: string;
}

export interface CellInfo {
  row: number;
  col: number;
  value: string;
  borderInfo?: BorderInfo;
  isMerged: boolean;
  mergedRange?: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  };
}