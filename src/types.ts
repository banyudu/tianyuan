export type 子目编号 = string; // 子目编号，如 1B-1, 7B-22

export interface CellData {
  row: number;
  col: number;
  address: string;
  value: any;
  type: string;
  merged: boolean;
  mergedRange?: {
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
  };
  borders: {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };
}

export interface ParsedExcelData {
  metadata: {
    filename: string;
    sheetName: string;
    totalRows: number;
    totalCols: number;
    actualRowCount: number;
    actualColCount: number;
    parsedAt: string;
  };
  cells: CellData[];
}

export interface 附注信息 {
  编号: 子目编号[]; // will be converted to a string with comma separated values
  附注信息: string;
}

export interface 附注信息行 {
  // row data of 附注信息表
  编号: string; // joined from 子目编号[]
  附注信息: string;
}

export interface 工作内容 {
  编号: 子目编号[]; // will be converted to a string with comma separated values
  工作内容: string;
}

export interface 工作内容行 {
  // row data of 工作内容表
  编号: string; // joined from 子目编号[]
  工作内容: string;
}

export interface 工作内容_附注信息表 {
  工作内容: 工作内容行[];
  附注信息: 附注信息行[];
}

// 子目信息表
export interface 子目章 {
  symbol: '$';
  编号: string;
  名称: string;
  children: 子目节[] | 子目项[];
}

export interface 子目节 {
  symbol: '$$';
  编号: string;
  名称: string;
  children: 子目小节[] | 子目项[];
}

export interface 子目小节 {
  symbol: '$$$';
  编号: string;
  名称: string;
  children: 子目小小节[] | 子目项[];
}

export interface 子目小小节 {
  symbol: '$$$$';
  编号: string;
  名称: string;
  children: 子目项[];
}

export interface 子目项 {
  定额号: 子目编号;
  子目名称: string;
  基价: number; // 0 if not found
  人工: number; // 0 if not found
  材料: number; // 0 if not found
  机械: number; // 0 if not found
  管理费: number; // 0 if not found
  利润: number; // 0 if not found
  其他: number; // 0 if not found
  图片名称?: string;
}

export type 子目行 = 子目章 | 子目节 | 子目小节 | 子目小小节 | 子目项;

export interface 子目信息表行 {
  // row data of 子目信息表
  symbol: '' | '$' | '$$' | '$$$' | '$$$$'; // 子目标题的符号，如果不是子目标题，则空字符串, symbol 列没有标题
  子目名称: string;
  基价: number; // 0 if not found
  人工: number; // 0 if not found
  材料: number; // 0 if not found
  机械: number; // 0 if not found
  管理费: number; // 0 if not found
  利润: number; // 0 if not found
  其他: number; // 0 if not found
  图片名称?: string;
}

export type 子目信息表 = 子目信息表行[];

export enum 材料类别 {
  // 1, 2, 3, 4, 5
  // TODO: confirm the meaning of 1, 2, 3, 4, 5
  主材 = 1,
  辅材 = 2,
  其他 = 3,
}

export interface 含量 {
  编号: 子目编号;
  名称: string; // example: 综合用工二类, 不锈钢电焊条奥102φ3.2, 聚四氟乙烯生料带宽20, 台式钻床钻孔直径16mm
  规格?: string;
  单位: string; // example: 工日, 台, m3, kg, 10个, 套, 台班
  单价: number;
  含量: number;
  主材标记: boolean; // 是否为主材
  材料号?: string;
  材料类别: 材料类别;
  是否有明细: boolean; // 是否有明细
}

export interface 材料表行 {
  编号: 子目编号;
  名称: string;
  规格?: string;
  单位: string; // example: 工日, 台, m3, kg, 10个, 套, 台班
  单价: number;
  含量: number;
  主材标记: boolean; // 是否为主材
  材料号?: string;
  材料类别: 材料类别;
  是否有明细: boolean; // 是否有明细
}
