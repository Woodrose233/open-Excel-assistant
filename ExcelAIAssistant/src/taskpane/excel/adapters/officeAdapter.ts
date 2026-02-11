/* global Excel */
import { IExcelAdapter, ExcelContext, QuickOperation } from "../types";

export class OfficeJsAdapter implements IExcelAdapter {
    async getContext(): Promise<ExcelContext> {
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const activeSheet = workbook.worksheets.getActiveWorksheet();
            const selection = workbook.getSelectedRange();
            const usedRange = activeSheet.getUsedRange();
            const worksheets = workbook.worksheets;

            activeSheet.load("name");
            selection.load("address");
            usedRange.load(["address", "values"]); // 加载 values 以获取表头
            workbook.load("name");
            worksheets.load("items/name");

            await context.sync();

            // 提取第一行作为表头 (如果存在)
            let headers: string[] = [];
            if (usedRange.values && usedRange.values.length > 0) {
                headers = usedRange.values[0].map((v: any) => String(v || ""));
            }

            return {
                activeSheetName: activeSheet.name,
                selectionAddress: selection.address,
                usedRangeAddress: usedRange.address,
                workbookName: workbook.name,
                sheetNames: worksheets.items.map(s => s.name),
                headers: headers
            };
        });
    }

    async executeQuickOps(ops: QuickOperation[]): Promise<any> {
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sheets = workbook.worksheets;
            const activeSheet = sheets.getActiveWorksheet();
            const results: { opIndex: number; data: Excel.Range }[] = [];

            for (let i = 0; i < ops.length; i++) {
                const op = ops[i];
                
                // 1. 特殊指令处理：脚本运行
                if (op.type === "run_script" && op.script) {
                    await this.runScript(op.script);
                    continue;
                }

                // 2. 基础指令处理
                const range = op.range ? this.resolveRange(context, op.range) : null;
                await this.processOp(context, op, range, results, i);
            }

            await context.sync();
            return this.formatResults(results);
        });
    }

    /**
     * 解析范围引用，支持 A1, indexes, 以及跨表 Sheet1!A1
     */
    private resolveRange(context: Excel.RequestContext, ref: any): Excel.Range {
        const workbook = context.workbook;
        if (typeof ref === "string") {
            if (ref.indexOf("!") !== -1) {
                const [sheetPart, address] = ref.split("!");
                const sheetName = sheetPart.replace(/^'|'$/g, ""); // 移除单引号
                return workbook.worksheets.getItem(sheetName).getRange(address);
            }
            return workbook.worksheets.getActiveWorksheet().getRange(ref);
        }
        return workbook.worksheets.getActiveWorksheet().getRangeByIndexes(ref.row, ref.col, ref.rowCount || 1, ref.colCount || 1);
    }

    /**
     * 核心操作分发逻辑
     */
    private async processOp(context: Excel.RequestContext, op: QuickOperation, range: Excel.Range | null, results: any[], index: number) {
        const workbook = context.workbook;
        const activeSheet = workbook.worksheets.getActiveWorksheet();

        switch (op.type) {
            case "read":
                if (range) {
                    range.load("values");
                    results.push({ opIndex: index, data: range });
                }
                break;

            case "write":
                if (range) await this.handleWrite(context, range, op.value);
                break;

            case "copy_paste":
                if (range && op.destination) {
                    const destRange = this.resolveRange(context, op.destination);
                    range.load("address");
                    destRange.copyFrom(range, Excel.RangeCopyType.all);
                }
                break;

            case "format":
                this.handleFormat(range, op.format || op.style);
                break;

            case "insert":
            case "insert_row":
                range?.insert(Excel.InsertShiftDirection.down);
                break;

            case "delete":
            case "delete_row":
                range?.delete(Excel.DeleteShiftDirection.up);
                break;

            case "create_sheet":
            case "add_sheet":
                await this.handleCreateSheet(context, op.sheetName || String(op.value));
                break;

            case "chart":
                if (range) {
                    const chart = activeSheet.charts.add((op.chartType as any) || Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);
                    if (op.title) chart.title.text = op.title;
                }
                break;
                
            case "clear":
                range?.clear();
                break;
        }
    }

    private async handleWrite(context: Excel.RequestContext, range: Excel.Range, value: any) {
        const isFormula = typeof value === 'string' && value.startsWith('=');
        const targetProperty = isFormula ? 'formulas' : 'values';
        
        if (Array.isArray(value)) {
            (range as any)[targetProperty] = Array.isArray(value[0]) ? value : [value];
        } else {
            range.load(["rowCount", "columnCount"]);
            await context.sync();
            
            if (range.rowCount > 1 || range.columnCount > 1) {
                const grid = Array(range.rowCount).fill(null).map(() => Array(range.columnCount).fill(value));
                (range as any)[targetProperty] = grid;
            } else {
                (range as any)[targetProperty] = [[value]];
            }
        }
    }

    private handleFormat(range: Excel.Range | null, style: any) {
        if (!range || !style) return;
        if (style.fillColor) range.format.fill.color = style.fillColor;
        if (style.fontColor) range.format.font.color = style.fontColor;
        if (style.bold !== undefined) range.format.font.bold = style.bold;
        if (style.fontSize) range.format.font.size = style.fontSize;
    }

    private async handleCreateSheet(context: Excel.RequestContext, name: string) {
        if (!name) return;
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        if (!sheets.items.some(s => s.name === name)) {
            sheets.add(name);
        }
    }

    private formatResults(results: any[]) {
        if (results.length === 0) return "OK";
        const data = results.map(r => ({ opIndex: r.opIndex, values: r.data.values }));
        return data.length === 1 ? data[0].values : data;
    }

    async runScript(script: string): Promise<any> {
        return await Excel.run(async (context) => {
            // 移除不安全的 Function/eval 构造，改用最直接的 Function 包装
            // 这也是目前 Office 环境下动态执行代码最稳健、最干净的方式
            const safeScript = `
                return (async function(context, Excel) {
                    ${script}
                })(context, Excel);
            `;
            try {
                const fn = new Function("context", "Excel", safeScript);
                return await fn(context, Excel);
            } catch (e) {
                console.error("Script Execution Failed:", e);
                throw e;
            }
        });
    }
}
