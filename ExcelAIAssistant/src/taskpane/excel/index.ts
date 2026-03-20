import { IExcelAdapter } from "./types";
import { OfficeJsAdapter } from "./adapters/officeAdapter";

/**
 * 平台工厂：根据环境自动选择适配器
 */
export function getExcelAdapter(): IExcelAdapter {
    if (typeof Office !== "undefined" && (window as any).Excel) {
        return new OfficeJsAdapter();
    }
    
    throw new Error("未检测到支持的电子表格环境 (Office)");
}
