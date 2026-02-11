export interface CellRange {
    row: number;
    col: number;
    rowCount?: number;
    colCount?: number;
    address?: string; // 如 "A1" 或 "Sheet1!A1:B2"
}

export type QuickOperationType = 
    | "read" 
    | "write" 
    | "format" 
    | "clear" 
    | "insert" 
    | "delete" 
    | "insert_row" 
    | "delete_row" 
    | "copy_paste" 
    | "create_sheet" 
    | "add_sheet" 
    | "run_script"
    | "chart";

export interface QuickOperation {
    type: QuickOperationType;
    range?: string | { row: number, col: number, rowCount?: number, colCount?: number };
    value?: any;
    format?: any;
    style?: any; // 兼容旧代码
    destination?: string | { row: number, col: number };
    sheetName?: string; // 用于 create_sheet 和 add_sheet
    script?: string; // 用于 run_script
    chartType?: string; // 用于 chart: ColumnClustered, Line, Pie 等
    title?: string; // 图表标题
}

export interface ExcelContext {
    activeSheetName: string;
    selectionAddress: string;
    usedRangeAddress: string;
    workbookName?: string;
    sheetNames?: string[];
    headers?: string[];
}

/**
 * 抽象适配器接口，用于隔离不同平台的 API 实现 (Office JS / WPS JS)
 */
export interface IExcelAdapter {
    /** 获取当前工作簿上下文信息 */
    getContext(): Promise<ExcelContext>;

    /** 执行一组原子快捷操作 */
    executeQuickOps(ops: QuickOperation[]): Promise<void>;

    /** 执行原生脚本 (在 Office 环境下是 JavaScript，WPS 下可能是其特有 API) */
    runScript(script: string): Promise<any>;
}
