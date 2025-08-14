import { Consumption } from "./types";

export interface TableRange {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface NormInfo {
  code: string;
  fullName?: string; // Complete name: ${baseName} ${specUnit} ${spec}&${unit}
  baseName?: string;
  specUnit?: string; // Unit specification like "风量(m3/h)"
  spec?: string;
  unit?: string;
  row: number;
  col: number;
  resources?: ResourceConsumption[]; // All resource consumptions for this norm
}

export interface ResourceConsumption {
  name: string;
  specification?: string; // 规格
  unit: string;
  consumption: string; // Keep as string to preserve trailing zeros
  isPrimary: boolean; // True if consumption was wrapped in parentheses
  category: string; // 人工/材料/机械
  categoryCode: number; // 1=人工, 2=材料, 3=机械, 5=other (5 for primary resources)
}

export interface ResourceInfo {
  category: string; // 人工/材料/机械 etc.
  names: string[]; // Multiple resource names from the same cell
  units: string[]; // Corresponding units for each name
  consumptions: Array<Record<string, Consumption>>; // consumption for each name-unit pair
  row: number;
}

export interface TableStructure {
  leadingElements?: {
    workContent?: string;
    unit?: string;
    row: number;
  };

  normCodesRow?: {
    labelCell: string; // "子目编号" etc.
    normCodes: NormInfo[];
    row: number;
  };

  normNamesRows?: {
    labelCell: string; // "子目名称" etc.
    normNames: Array<{
      baseName: string;
      specUnit?: string; // Unit specification like "风量(m3/h)"
      spec?: string;
      unit?: string;
      fullName: string; // ${baseName} ${specUnit} ${spec}&${unit}
      normCode: string; // corresponding norm code
      col: number;
    }>;
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
  normCodes: string[];
  unit?: string;
  workContent?: string;  // Optional - not every table has work content
  notes: string[];       // Optional - not every table has notes
  isContinuation?: boolean;
  continuationOf?: string;

  // New detailed structure
  structure?: TableStructure;
  norms?: NormInfo[]; // All norms in this table with their resource consumptions
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
