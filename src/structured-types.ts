export interface TableRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface QuotaCodeInfo {
  code: string;
  name?: string;
  baseName?: string;
  spec?: string;
  unit?: string;
  row: number;
  col: number;
}

export interface ResourceInfo {
  category: string; // 人工/材料/机械 etc.
  name: string;
  unit?: string;
  consumptions: { [quotaCode: string]: number | string }; // quota code -> consumption amount
  row: number;
}

export interface TableStructure {
  leadingElements?: {
    workContent?: string;
    unit?: string;
    row: number;
  };
  
  quotaCodesRow?: {
    labelCell: string; // "子目编号" etc.
    quotaCodes: QuotaCodeInfo[];
    row: number;
  };
  
  quotaNamesRows?: {
    labelCell: string; // "子目名称" etc.
    baseNames: string[]; // first row after label
    specs: string[];     // second row after label (if exists)
    startRow: number;
    endRow: number;
  };
  
  resourcesSection?: {
    labelCell: string; // "人材机名称" etc.
    unitLabelCell?: string; // "单位" 
    consumptionLabelCell?: string; // "消耗量"
    resources: ResourceInfo[];
    startRow: number;
    endRow: number;
  };
  
  trailingElements?: {
    notes: string[];
    rows: number[];
  };
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
  
  // New detailed structure
  structure?: TableStructure;
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
  description?: string[]; // Optional description text
  subSections: SubSection[];
  tableAreas: TableArea[];
}

export interface Chapter {
  id: string;
  name: string;
  number: string;
  description?: string[]; // Optional description text
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