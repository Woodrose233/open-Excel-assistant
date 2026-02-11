import { IExcelAdapter } from "./types";
import { OfficeJsAdapter } from "./adapters/officeAdapter";

/**
 * 平台工厂：根据环境自动选择适配器
 * 后续移植到 WPS 时，只需在此增加判断逻辑并提供 WpsAdapter 即可
 */
export function getExcelAdapter(): IExcelAdapter {
    // 简单的环境检测逻辑
    if (typeof Office !== "undefined" && (window as any).Excel) {
        return new OfficeJsAdapter();
    }
    
    // 如果是 WPS 环境，可以在此返回 WpsAdapter
    // if (typeof wps !== "undefined") { return new WpsAdapter(); }
    
    throw new Error("未检测到支持的电子表格环境 (Office/WPS)");
}
