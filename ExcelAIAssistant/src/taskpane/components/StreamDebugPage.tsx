import * as React from "react";
import { callAI, Message, AIConfig } from "../utils/api";
import { ApiConfig } from "./types";

interface StreamDebugPageProps {
  config: ApiConfig | null;
}

const StreamDebugPage: React.FC<StreamDebugPageProps> = ({ config }) => {
  const [prompt, setPrompt] = React.useState("请写一段关于 Excel 自动化的短文，大约 200 字，要求分成三个段落。");
  const [isStreaming, setIsStreaming] = React.useState(false);
  const [streamContent, setStreamContent] = React.useState("");
  const [debugLogs, setDebugLogs] = React.useState<string[]>([]);
  const [rawChunks, setRawChunks] = React.useState<string[]>([]);

  const addLog = (msg: string) => {
    setDebugLogs(prev => [`[${new Date().toLocaleTimeString()}] ${msg}`, ...prev.slice(0, 49)]);
  };

  const testStream = async () => {
    if (!config) return;
    
    setIsStreaming(true);
    setStreamContent("");
    setDebugLogs([]);
    setRawChunks([]);
    addLog("开始流式测试...");

    const messages: Message[] = [
      { role: "system", content: "你是一个专业的助手。" },
      { role: "user", content: prompt }
    ];

    try {
      const { content: fullContent } = await callAI(
        messages,
        config as unknown as AIConfig,
        (delta) => {
          setStreamContent(prev => prev + delta);
          // 记录非空的 delta
          if (delta.trim()) {
            addLog(`收到 Delta: "${delta.substring(0, 20)}${delta.length > 20 ? '...' : ''}"`);
          }
        }
      );
      addLog("流式请求完成。");
      addLog(`最终内容长度: ${fullContent.length}`);
    } catch (e) {
      addLog(`❌ 错误: ${String(e)}`);
    } finally {
      setIsStreaming(false);
    }
  };

  return (
    <div style={{ padding: "15px", display: "flex", flexDirection: "column", gap: "10px", height: "100%", overflow: "hidden" }}>
      <h3 style={{ margin: "0" }}>流式输出调试工具</h3>
      
      <div style={{ fontSize: "12px", color: "#666" }}>
        当前配置: {config?.type} | 模型: {config?.model}
      </div>

      <textarea
        value={prompt}
        onChange={(e) => setPrompt(e.target.value)}
        style={{ width: "100%", height: "60px", padding: "8px", boxSizing: "border-box", fontSize: "12px" }}
      />

      <button 
        onClick={testStream} 
        disabled={isStreaming || !config}
        style={{
          padding: "10px",
          backgroundColor: isStreaming ? "#ccc" : "#0078d4",
          color: "white",
          border: "none",
          borderRadius: "4px",
          cursor: isStreaming ? "not-allowed" : "pointer"
        }}
      >
        {isStreaming ? "正在输出..." : "开始测试流式输出"}
      </button>

      <div style={{ display: "flex", flex: 1, gap: "10px", overflow: "hidden" }}>
        {/* 左侧：实时内容 */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
          <div style={{ fontSize: "11px", fontWeight: "bold", marginBottom: "4px" }}>实时内容预览:</div>
          <div style={{ 
            flex: 1, 
            padding: "8px", 
            backgroundColor: "#fff", 
            border: "1px solid #ddd", 
            borderRadius: "4px",
            overflowY: "auto",
            fontSize: "13px",
            lineHeight: "1.5",
            whiteSpace: "pre-wrap"
          }}>
            {streamContent || (isStreaming ? "等待首字节..." : "暂无内容")}
          </div>
        </div>

        {/* 右侧：调试日志 */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
          <div style={{ fontSize: "11px", fontWeight: "bold", marginBottom: "4px" }}>调试日志:</div>
          <div style={{ 
            flex: 1, 
            padding: "8px", 
            backgroundColor: "#000", 
            color: "#0f0", 
            borderRadius: "4px",
            overflowY: "auto",
            fontSize: "10px",
            fontFamily: "monospace"
          }}>
            {debugLogs.length === 0 ? "等待运行..." : debugLogs.map((log, i) => (
              <div key={i} style={{ marginBottom: "2px" }}>{log}</div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

export default StreamDebugPage;
