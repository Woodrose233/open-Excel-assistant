import * as React from "react";
import { getExcelAdapter } from "../excel";
import { QuickOperation } from "../excel/types";

const DebugTestPage: React.FC = () => {
  const [result, setResult] = React.useState<string>("");
  const excelAdapter = React.useMemo(() => getExcelAdapter(), []);

  const testQuickOp = async () => {
    try {
      setResult("正在测试快捷指令...");
      const op: QuickOperation = {
        type: "write",
        range: "A1",
        value: "测试成功: " + new Date().toLocaleTimeString()
      };
      await excelAdapter.executeQuickOps([op]);
      setResult("快捷指令执行成功！查看单元格 A1");
    } catch (e) {
      setResult("快捷指令失败: " + String(e));
    }
  };

  const testScript = async () => {
    try {
      setResult("正在测试脚本...");
      const script = `
        Excel.run(function (context) {
          var sheet = context.workbook.worksheets.getActiveWorksheet();
          var range = sheet.getRange("A2");
          range.values = [["脚本执行成功: " + new Date().toLocaleTimeString()]];
          range.format.fill.color = "yellow";
          return context.sync();
        });
      `;
      // @ts-ignore
      if (excelAdapter.runScript) {
        // @ts-ignore
        await excelAdapter.runScript(script);
        setResult("脚本执行成功！查看单元格 A2");
      } else {
        // 如果没有 runScript，尝试通过 executeQuickOps 包装
        const op: QuickOperation = {
          type: "run_script",
          script: script
        };
        await excelAdapter.executeQuickOps([op]);
        setResult("脚本通过 QuickOp 包装执行成功！查看单元格 A2");
      }
    } catch (e) {
      setResult("脚本失败: " + String(e));
    }
  };

  const testGetContext = async () => {
    try {
      setResult("正在获取上下文...");
      const context = await excelAdapter.getContext();
      setResult("获取上下文成功:\n" + JSON.stringify(context, null, 2));
    } catch (e) {
      setResult("获取上下文失败: " + String(e));
    }
  };

  const testCopyPaste = async () => {
    try {
      setResult("正在测试复制粘贴...\n1. 写入源数据 (A1:B2)");
      // 先写点多行多列的数据
      await excelAdapter.executeQuickOps([
        { type: "write", range: "A1", value: "R1C1" },
        { type: "write", range: "B1", value: "R1C2" },
        { type: "write", range: "A2", value: "R2C1" },
        { type: "write", range: "B2", value: "R2C2" },
        { type: "format", range: "A1:B2", style: { fillColor: "pink", bold: true } }
      ]);
      
      setResult(prev => prev + "\n2. 执行复制粘贴 (A1:B2 -> D1)");
      // 执行复制粘贴
      await excelAdapter.executeQuickOps([
        { type: "copy_paste", range: "A1:B2", destination: "D1" }
      ]);
      
      setResult(prev => prev + "\n3. 验证结果...");
      const readResult = await excelAdapter.executeQuickOps([
        { type: "read", range: "D1:E2" }
      ]);
      
      setResult(prev => prev + "\n验证数据: " + JSON.stringify(readResult));
      setResult(prev => prev + "\n\n✅ 复制粘贴测试完成！请检查 D1:E2 是否有粉色背景和对应文字。");
    } catch (e) {
      setResult("复制粘贴失败: " + String(e));
    }
  };

  const cleanTestArea = async () => {
    try {
      setResult("正在清空测试区域 (A1:Z100)...");
      await excelAdapter.executeQuickOps([
        { type: "clear", range: "A1:Z100" }
      ]);
      setResult("区域已清空");
    } catch (e) {
      setResult("清空失败: " + String(e));
    }
  };

  const testCrossSheetCopy = async () => {
        try {
            setResult("正在测试跨表复制粘贴...\n1. 检查或创建目标表 (Sheet2)");
            // 使用快捷指令创建表，避开 runScript 的语法问题
            await excelAdapter.executeQuickOps([
                { type: "create_sheet", sheetName: "Sheet2" }
            ]);

            setResult(prev => prev + "\n2. 写入源数据到当前表 (A1)");
            await excelAdapter.executeQuickOps([
                { type: "write", range: "A1", value: "CrossSheetData" }
            ]);

            setResult(prev => prev + "\n3. 执行跨表复制 (当前表!A1 -> Sheet2!B1)");
            const contextData = await excelAdapter.getContext();
            // 使用带引号的表名进行更严格的测试
            const sourceRef = `'${contextData.activeSheetName}'!A1`;
            const destRef = `Sheet2!B1`;
            
            await excelAdapter.executeQuickOps([
                { type: "copy_paste", range: sourceRef, destination: destRef }
            ]);

            setResult(prev => prev + "\n4. 验证 Sheet2 结果...");
            const readResult = await excelAdapter.executeQuickOps([
                { type: "read", range: "Sheet2!B1" }
            ]);

            setResult(prev => prev + "\n验证数据: " + JSON.stringify(readResult));
            setResult(prev => prev + "\n\n✅ 跨表复制测试完成！请检查 Sheet2!B1 是否为 'CrossSheetData'。");
        } catch (e) {
            setResult("跨表复制失败: " + String(e));
        }
    };

  return (
    <div style={{ padding: "15px", display: "flex", flexDirection: "column", gap: "10px" }}>
      <h3 style={{ margin: "0 0 10px 0" }}>功能调试页</h3>
      
      <button onClick={testQuickOp} style={buttonStyle}>测试快捷指令 (A1)</button>
      <button onClick={testScript} style={buttonStyle}>测试脚本执行 (A2)</button>
      <button onClick={testGetContext} style={buttonStyle}>测试获取上下文</button>
      <button onClick={testCopyPaste} style={buttonStyle}>测试复制粘贴 (A1:B2 ➜ D1)</button>
      <button onClick={testCrossSheetCopy} style={{ ...buttonStyle, backgroundColor: "#e1f5fe", color: "#000" }}>测试跨表复制 (当前!A1 ➜ Sheet2!B1)</button>
      <button onClick={cleanTestArea} style={{ ...buttonStyle, backgroundColor: "#d83b01" }}>清空测试区域</button>
      
      <div style={{ 
        marginTop: "10px", 
        padding: "10px", 
        backgroundColor: "#f5f5f5", 
        borderRadius: "4px",
        fontSize: "12px",
        whiteSpace: "pre-wrap",
        border: "1px solid #ddd",
        minHeight: "100px"
      }}>
        {result || "等待操作..."}
      </div>
    </div>
  );
};

const buttonStyle: React.CSSProperties = {
  padding: "8px",
  backgroundColor: "#0078d4",
  color: "white",
  border: "none",
  borderRadius: "4px",
  cursor: "pointer"
};

export default DebugTestPage;
