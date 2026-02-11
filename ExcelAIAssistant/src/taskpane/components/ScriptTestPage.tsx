import * as React from "react";
import { getExcelAdapter } from "../excel";

const ScriptTestPage: React.FC = () => {
  const [script, setScript] = React.useState(`const sheet = context.workbook.worksheets.getActiveWorksheet();
sheet.getRange("A1").values = [["分步测试1"]];
await context.sync();`);
  const [result, setResult] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const excelAdapter = React.useMemo(() => getExcelAdapter(), []);

  const runTest = async (code: string) => {
    setResult("");
    setError("");
    try {
      const res = await excelAdapter.runScript(code);
      setResult(JSON.stringify(res) || "执行成功");
    } catch (e: any) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div style={{ padding: "10px", display: "flex", flexDirection: "column", gap: "10px" }}>
      <h3 style={{ margin: 0, fontSize: "14px" }}>脚本分步调试</h3>
      
      <div style={{ display: "flex", flexWrap: "wrap", gap: "5px" }}>
        <button onClick={() => runTest('resolve("Hello World");')} style={btnStyle}>测试1: 基础返回</button>
        <button onClick={() => runTest('var s = context.workbook.worksheets.getActiveWorksheet(); s.load("name"); context.sync().then(function() { resolve("当前表: " + s.name); });')} style={btnStyle}>测试2: API调用</button>
        <button onClick={() => runTest('context.workbook.worksheets.getActiveWorksheet().getRange("A1").values = [["测试"]]; context.sync().then(function() { resolve("写入成功"); });')} style={btnStyle}>测试3: 写入A1</button>
      </div>

      <textarea
        value={script}
        onChange={(e) => setScript(e.target.value)}
        style={{
          width: "100%",
          height: "100px",
          fontFamily: "monospace",
          fontSize: "12px",
          padding: "5px",
          boxSizing: "border-box"
        }}
      />
      
      <button onClick={() => runTest(script)} style={{ ...btnStyle, backgroundColor: "#217346" }}>运行输入框代码</button>
      
      {result && <div style={{ padding: "8px", backgroundColor: "#e6f4ea", color: "#137333", fontSize: "12px" }}><b>结果:</b> {result}</div>}
      {error && <div style={{ padding: "8px", backgroundColor: "#fce8e6", color: "#c5221f", fontSize: "12px" }}><b>报错:</b> {error}</div>}
    </div>
  );
};

const btnStyle: React.CSSProperties = {
  padding: "6px 10px",
  backgroundColor: "#f3f2f1",
  border: "1px solid #8a8886",
  borderRadius: "2px",
  cursor: "pointer",
  fontSize: "12px"
};

export default ScriptTestPage;
