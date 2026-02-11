import * as React from "react";
import { ApiConfig } from "./types";
import { callAI, Message, SYSTEM_PROMPT, compressContext, AIConfig } from "../utils/api";
import { getExcelAdapter } from "../excel";
import { QuickOperation } from "../excel/types";

interface ChatPageProps {
  config: ApiConfig | null;
  theme: {
    isDark: boolean;
    primaryColor: string;
    bgColor: string;
    textColor: string;
    borderColor: string;
  };
}

export type { Message };

// 1. 抽取通用的 Markdown 渲染与卡片展示逻辑
const MessageBubble = React.memo(({ 
  msg, 
  theme, 
  onExecuteAction, 
  onRunScript 
}: { 
  msg: Message, 
  theme: any, 
  onExecuteAction: (content: string) => void,
  onRunScript: (script: string) => void
}) => {
  const { content, role } = msg;

  // 预处理：如果 AI 忘记加代码块但输出了纯 JSON，尝试补全它
  let processedText = content.trim();
  if (processedText.startsWith('{') && processedText.endsWith('}') && !processedText.includes('```')) {
    try {
      JSON.parse(processedText);
      processedText = `\`\`\`json\n${processedText}\n\`\`\``;
    } catch (e) {}
  }

  // 处理代码块 (JSON/Javascript)，渲染为卡片
  const parts = processedText.split(/(```(?:json|javascript|js)[\s\S]*?```)/g);

  const getOpName = (type: string) => {
    const names: Record<string, string> = {
      "read": "读取数据",
      "write": "写入数据",
      "format": "格式设置",
      "clear": "清除内容",
      "insert": "插入行/列",
      "delete": "删除行/列",
      "add_sheet": "新建工作表",
      "create_sheet": "新建工作表",
      "run_script": "运行脚本",
      "copy_paste": "复制粘贴",
      "chart": "绘制图表"
    };
    return names[type] || "未知操作";
  };

  return (
    <div 
      style={{ 
        maxWidth: "85%",
        padding: "8px 12px",
        borderRadius: "8px",
        backgroundColor: role === "user" ? theme.primaryColor : (theme.isDark ? "#444444" : "white"),
        color: role === "user" ? "#ffffff" : theme.textColor,
        boxShadow: theme.isDark ? "0 2px 4px rgba(0,0,0,0.3)" : "0 1px 2px rgba(0,0,0,0.2)",
        border: role === "assistant" ? `1px solid ${theme.borderColor}` : "none",
        whiteSpace: "pre-wrap",
        wordBreak: "break-word",
        overflowWrap: "anywhere",
        fontSize: "13px",
        lineHeight: "1.5",
        boxSizing: "border-box"
      }}
    >
      {parts.map((part, i) => {
        const jsonMatch = part.match(/```json\s*([\s\S]*?)\s*```/);
        const scriptMatch = part.match(/```(?:javascript|js)\s*([\s\S]*?)\s*```/);

        if (jsonMatch && role === "assistant") {
          let ops: QuickOperation[] = [];
          try {
            const parsed = JSON.parse(jsonMatch[1]);
            ops = Array.isArray(parsed) ? parsed : [parsed];
          } catch (e) {}

          const opNames = Array.from(new Set(ops.map(op => getOpName(op.type)))).join(", ");

          return (
            <div key={i} style={{ 
              marginTop: "8px", 
              padding: "8px", 
              backgroundColor: theme.isDark ? "#333333" : "#f0f0f0", 
              borderRadius: "4px",
              borderLeft: `4px solid ${theme.primaryColor}`
            }}>
              <div style={{ fontSize: "11px", fontWeight: "bold", marginBottom: "4px", color: theme.textColor }}>
                快捷指令: {opNames || "解析中..."}
              </div>
              <button 
                onClick={() => onExecuteAction(part)}
                style={{
                  width: "100%",
                  padding: "6px",
                  backgroundColor: theme.primaryColor,
                  color: "white",
                  border: "none",
                  borderRadius: "2px",
                  cursor: "pointer",
                  fontSize: "12px"
                }}
              >
                执行操作
              </button>
            </div>
          );
        }

        if (scriptMatch && role === "assistant") {
          return (
            <div key={i} style={{ 
              marginTop: "8px", 
              padding: "8px", 
              backgroundColor: theme.isDark ? "#333333" : "#f0f0f0", 
              borderRadius: "4px",
              borderLeft: "4px solid #f1c40f"
            }}>
              <div style={{ fontSize: "11px", fontWeight: "bold", marginBottom: "4px", color: theme.textColor }}>脚本逻辑</div>
              <code style={{ fontSize: "10px", display: "block", marginBottom: "4px", overflowX: "auto", color: theme.textColor }}>
                {scriptMatch[1].length > 60 ? scriptMatch[1].substring(0, 60) + "..." : scriptMatch[1]}
              </code>
              <button 
                onClick={() => onRunScript(scriptMatch[1])}
                style={{
                  width: "100%",
                  padding: "6px",
                  backgroundColor: "#f1c40f",
                  color: "black",
                  border: "none",
                  borderRadius: "2px",
                  cursor: "pointer",
                  fontSize: "12px"
                }}
              >
                运行脚本
              </button>
            </div>
          );
        }

        // 简单的 MD 解析
        let html = part
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/^### (.*$)/gim, '<h3 style="margin:8px 0;font-size:15px;">$1</h3>')
          .replace(/^## (.*$)/gim, '<h2 style="margin:10px 0;font-size:17px;">$1</h2>')
          .replace(/^# (.*$)/gim, '<h1 style="margin:12px 0;font-size:19px;">$1</h1>')
          .replace(/\*\*(.*)\*\*/gim, '<strong>$1</strong>')
          .replace(/\*(.*)\*/gim, '<em>$1</em>')
          .replace(/^\- (.*$)/gim, '<li style="margin-left:15px;">$1</li>');
        
        return <span key={i} dangerouslySetInnerHTML={{ __html: html }} />;
      })}
    </div>
  );
});

const ChatPage: React.FC<ChatPageProps> = ({ config, theme }) => {
  const [messages, setMessages] = React.useState<Message[]>([]);
  const [input, setInput] = React.useState("");
  const [isLoading, setIsLoading] = React.useState(false);
  const [lastResult, setLastResult] = React.useState<string | null>(null); 
  const [pendingFeedback, setPendingFeedback] = React.useState<string | null>(null); // 新增：待自动反馈的结果
  const [iterationCount, setIterationCount] = React.useState(0); // 新增：迭代计数器，防止无限循环
  const [error, setError] = React.useState<string | null>(null);
  const [clearFeedback, setClearFeedback] = React.useState<string | null>(null);
  const [showConfirm, setShowConfirm] = React.useState(false);
  const isComposing = React.useRef(false);
  const isInitialized = React.useRef(false);
  const abortControllerRef = React.useRef<{ abort: () => void } | null>(null);
  const messagesEndRef = React.useRef<HTMLDivElement>(null);
  
  const excelAdapter = React.useMemo(() => getExcelAdapter(), []);

  const scrollToBottom = () => {
    // 使用 requestAnimationFrame 确保在 DOM 更新后滚动
    requestAnimationFrame(() => {
      messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
    });
  };

  // 1. 初始化加载
  React.useEffect(() => {
    const saved = localStorage.getItem("chat_history");
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setMessages(parsed);
        }
      } catch (e) {
        console.error("加载对话历史失败", e);
      }
    }
    isInitialized.current = true;

    // “土办法”：首次加载后模拟一次宽度微调，强制浏览器重绘布局
    // 这种方法在 Office 侧边栏渲染引擎出现奇怪的横向溢出残留时非常有效
    setTimeout(() => {
      const container = document.body;
      if (container) {
        const originalWidth = container.style.width;
        // 缩窄 1px 再恢复，触发 Layout Engine 重新计算
        container.style.width = "calc(100% - 1px)";
        
        requestAnimationFrame(() => {
          container.style.width = originalWidth;
          // 同时手动触发一次 resize 事件，通知所有监听者
          window.dispatchEvent(new Event("resize"));
        });
      }
    }, 300); // 稍微延迟，等待消息气泡初步渲染完成
  }, []);

  // 2. 对话持久化保存和滚动
  React.useEffect(() => {
    if (!isInitialized.current) return;

    if (messages.length > 0) {
      localStorage.setItem("chat_history", JSON.stringify(messages));
    } else {
      localStorage.removeItem("chat_history");
    }
    scrollToBottom();
  }, [messages]);

  // 3. 自动反馈循环逻辑
  React.useEffect(() => {
    if (pendingFeedback && !isLoading) {
      const feedback = pendingFeedback;
      setPendingFeedback(null);
      // 延迟一小会儿，让用户看清 AI 的上一条输出
      setTimeout(() => {
        handleSend(feedback, true);
      }, 800);
    }
  }, [pendingFeedback, isLoading]);

  const clearHistory = () => {
    setShowConfirm(true);
  };

  const confirmClear = () => {
    try {
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }
      
      localStorage.removeItem("chat_history");
      setMessages([]);
      setInput(""); // 清空输入框
      setError(null); // 清空错误
      setIterationCount(0); // 重置计数器
      
      setClearFeedback("对话已清空");
      setShowConfirm(false);
      setTimeout(() => setClearFeedback(null), 3000);
    } catch (err: any) {
      setError(`清除失败: ${err instanceof Error ? err.message : String(err)}`);
      setShowConfirm(false);
    }
  };

  const checkAndExecuteAction = React.useCallback(async (content: string) => {
    // 1. 提取 QuickOps JSON
    const jsonMatch = content.match(/```json\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      try {
        const ops = JSON.parse(jsonMatch[1]);
        const opsArray = Array.isArray(ops) ? ops : [ops];
        const result = await excelAdapter.executeQuickOps(opsArray);
        console.log("指令执行成功", result);
        return;
      } catch (e) {
        setError(`指令执行失败: ${e instanceof Error ? e.message : String(e)}`);
      }
    }

    // 2. 提取 JavaScript 脚本
    const scriptMatch = content.match(/```(?:javascript|js)\s*([\s\S]*?)\s*```/);
    if (scriptMatch) {
      try {
        const result = await excelAdapter.runScript(scriptMatch[1]);
        console.log("脚本执行成功", result);
        return;
      } catch (e) {
        setError(`脚本执行失败: ${e instanceof Error ? e.message : String(e)}`);
      }
    }
  }, [excelAdapter]);

  const handleRunScript = React.useCallback(async (script: string) => {
    setIsLoading(true);
    setError(null);
    try {
      const scriptOp: QuickOperation = {
        type: "run_script",
        script: script
      };
      await excelAdapter.executeQuickOps([scriptOp]);
    } catch (e: any) {
      console.error("脚本执行失败", e);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsLoading(false);
    }
  }, [excelAdapter]);

  const handleSendIteration = async () => {
    // 该功能已按用户要求停用
  };

  const handleSend = async (customInput?: string, isAutoFeedback: boolean = false) => {
    if (!config || isLoading) return;
    
    const textToSend = typeof customInput === "string" ? customInput : input;
    if (!textToSend.trim() && !isAutoFeedback && typeof customInput !== "string") return;

    // 立即清空输入框，防止重复发送
    if (typeof customInput !== "string") setInput("");
    
    // 更新迭代计数
    let currentIteration = 0;
    if (isAutoFeedback) {
      currentIteration = iterationCount + 1;
      setIterationCount(currentIteration);
      
      // 强制停止机制：超过 10 轮自动操作
      if (currentIteration > 10) {
        setError("已达到连续操作上限（10轮），为防止死循环已强制停止。您可以继续输入指令。");
        setIsLoading(false);
        setIterationCount(0); // 停止后重置，允许用户手动继续
        return;
      }
    } else {
      setIterationCount(0);
    }

    setIsLoading(true);
    setError(null);
    let finalContent = "";

    try {
      // 1. 获取环境信息
      let envInfo = "";
      try {
        const context = await excelAdapter.getContext();
        const headerStr = context.headers && context.headers.length > 0 ? ` [表头] ${context.headers.join(", ")}` : "";
        envInfo = `[当前环境] 文件: ${context.workbookName || "未知"}; 工作表: ${context.activeSheetName}; 选区: ${context.selectionAddress}; 数据范围: ${context.usedRangeAddress}${headerStr}`;
      } catch (e) {
        console.error("Failed to get context:", e);
      }

      // 2. 构造并更新消息列表
      const newUserMessage: Message = { 
        role: "user", 
        content: isAutoFeedback 
          ? `${envInfo}\n\n[系统自动反馈] ${textToSend}` 
          : (envInfo ? `${envInfo}\n\n${textToSend}` : textToSend) 
      };

      // @ts-ignore
      if (isAutoFeedback) newUserMessage.isHidden = true;

      // 同步构造下一轮完整的消息历史
      const nextMessages = [...messages];
      if (!isAutoFeedback && lastResult) {
        nextMessages.push({
          role: "user",
          content: `[上一轮操作执行结果] ${lastResult}`,
          // @ts-ignore
          isHidden: true
        });
        setLastResult(null);
      }
      nextMessages.push(newUserMessage);
      
      // 更新 UI 状态
      setMessages(nextMessages);

      // 3. 获取当前选区上下文并准备发送给 AI
      let excelContext = null;
      try {
        excelContext = await excelAdapter.getContext();
      } catch (e) {
        console.warn("获取上下文失败", e);
      }

      const systemContent = SYSTEM_PROMPT.replace("{context}", JSON.stringify(excelContext || {}, null, 2));
      
      const historyMessages = compressContext(nextMessages, 20).map(m => ({
        role: m.role,
        content: m.content
      }));

      const apiMessages = [
        { role: "system", content: systemContent },
        ...historyMessages
      ];

      // 创建 AI 消息占位
      setMessages(prev => [...prev, { role: "assistant", content: "" }]);

      const { content: fullContent, abort } = await callAI(
        apiMessages, 
        config,
        (delta) => {
          setMessages(prev => {
            const next = [...prev];
            const lastIdx = next.length - 1;
            if (lastIdx >= 0 && next[lastIdx].role === "assistant") {
              // 重要：必须创建新对象副本，否则 React.memo 会认为 props 没变而不触发重绘
              next[lastIdx] = { 
                ...next[lastIdx], 
                content: next[lastIdx].content + delta 
              };
            }
            return next;
          });
        }
      );

      abortControllerRef.current = { abort };
      finalContent = fullContent;

      // 4. 自动顺序执行所有提取到的操作 (JSON 或 脚本)
      const allBlocks = finalContent.match(/```(?:json|javascript|js)[\s\S]*?```/g);
      const isFinished = finalContent.includes("[FINISH]");

      if (allBlocks && allBlocks.length > 0) {
        let executionSummary = "";
        
        for (const block of allBlocks) {
          const jsonMatch = block.match(/```json\s*([\s\S]*?)\s*```/);
          const scriptMatch = block.match(/```(?:javascript|js)\s*([\s\S]*?)\s*```/);

          if (jsonMatch) {
            try {
              const ops = JSON.parse(jsonMatch[1]);
              const opsArray = Array.isArray(ops) ? ops : [ops];
              const result = await excelAdapter.executeQuickOps(opsArray);
              executionSummary += `[快捷指令] 执行成功: ${JSON.stringify(result)}; `;
            } catch (e) {
              executionSummary += `[快捷指令] 执行失败: ${e instanceof Error ? e.message : String(e)}; `;
            }
          }

          if (scriptMatch) {
            try {
              const result = await excelAdapter.runScript(scriptMatch[1]);
              executionSummary += `[脚本] 执行成功: ${JSON.stringify(result)}; `;
            } catch (e) {
              executionSummary += `[脚本] 执行失败: ${e instanceof Error ? e.message : String(e)}; `;
            }
          }
        }

        if (executionSummary && !isFinished) {
          setPendingFeedback(executionSummary);
          console.log("所有操作已自动执行，已排队自动反馈:", executionSummary);
        } else if (executionSummary && isFinished) {
          console.log("所有操作已执行，但 AI 标记了 [FINISH]，停止自动反馈循环。");
          setLastResult(executionSummary); // 仍然存入 lastResult 供下次手动提问使用
        }
      }

    } catch (e: any) {
      if (e.name === "AbortError" || e.message === "AbortError" || e.message === "cancelled") {
        console.log("请求已取消");
      } else {
        setError(e.message || "请求失败，请稍后重试。");
      }
    } finally {
      setIsLoading(false);
      abortControllerRef.current = null;
    }
  };

  const handleStop = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      setIsLoading(false);
      abortControllerRef.current = null;
      setIterationCount(0); // 用户手动停止，重置迭代计数
      setPendingFeedback(null); // 清除排队的反馈
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    // 使用 ref 追踪组合输入状态，这是最稳妥的 IME 处理方式
    if (isComposing.current) {
      return;
    }
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  if (!config) {
    return (
      <div style={{ textAlign: "center", marginTop: "30px", color: theme.isDark ? "#bbbbbb" : "#605e5c", fontSize: "13px" }}>
        <p>请先在“设置”中配置并选择一个 API。</p>
      </div>
    );
  }

  return (
    <div style={{ 
      display: "flex", 
      flexDirection: "column", 
      height: "100%", 
      overflow: "hidden", // 顶层容器禁止滚动，由内部 Content 处理
      boxSizing: "border-box"
    }}>
      {/* 1. Header: 固定高度 */}
      <div style={{ 
        display: "flex", 
        justifyContent: "space-between", 
        alignItems: "center",
        padding: "4px 8px", 
        borderBottom: `1px solid ${theme.borderColor}`,
        backgroundColor: theme.isDark ? "#333333" : "#f3f2f1",
        flexShrink: 0,
        height: "32px",
        boxSizing: "border-box"
      }}>
        <div style={{ fontSize: "11px", color: "#40bc74", fontWeight: "bold" }}>
          {clearFeedback && `✓ ${clearFeedback}`}
        </div>
        {!showConfirm ? (
          <button 
            onClick={clearHistory}
            style={{
              background: "none",
              border: "none",
              color: theme.isDark ? "#aaaaaa" : "#666666",
              cursor: "pointer",
              fontSize: "11px"
            }}
          >
            🗑️ 清除历史
          </button>
        ) : (
          <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
            <span style={{ fontSize: "11px", color: theme.textColor }}>确定清除？</span>
            <button 
              onClick={confirmClear}
              style={{
                padding: "2px 8px",
                backgroundColor: "#d83b01",
                color: "white",
                border: "none",
                borderRadius: "2px",
                cursor: "pointer",
                fontSize: "10px"
              }}
            >
              确定
            </button>
            <button 
              onClick={() => setShowConfirm(false)}
              style={{
                padding: "2px 8px",
                backgroundColor: theme.isDark ? "#555555" : "#e1dfdd",
                color: theme.textColor,
                border: "none",
                borderRadius: "2px",
                cursor: "pointer",
                fontSize: "10px"
              }}
            >
              取消
            </button>
          </div>
        )}
      </div>

      {/* 2. Content: 自动伸缩，处理滚动 */}
      <div style={{ 
        flex: 1, 
        overflowY: "auto", 
        overflowX: "hidden",
        padding: "8px", 
        display: "flex", 
        flexDirection: "column", 
        gap: "8px",
        boxSizing: "border-box"
      }}>
        {messages.length === 0 && (
          <div style={{ textAlign: "center", color: theme.isDark ? "#bbbbbb" : "#999999", marginTop: "15px", fontSize: "13px" }}>
            你好！我是你的 Excel AI 助手。
          </div>
        )}
        {messages.filter(m => !(m as any).isHidden).map((msg, index) => (
          <div 
            key={index} 
            style={{ 
              display: "flex",
              flexDirection: "column",
              alignItems: msg.role === "user" ? "flex-end" : "flex-start",
              width: "100%",
              flexShrink: 0,
              boxSizing: "border-box"
            }}
          >
            <MessageBubble 
              msg={msg} 
              theme={theme} 
              onExecuteAction={checkAndExecuteAction}
              onRunScript={handleRunScript}
            />
          </div>
        ))}
        {isLoading && (
          <div style={{ alignSelf: "flex-start", padding: "8px", color: theme.isDark ? "#aaaaaa" : "#605e5c", fontStyle: "italic", fontSize: "12px" }}>
            AI 正在思考中...
          </div>
        )}
        {error && (
          <div style={{ 
            alignSelf: "center", 
            padding: "8px", 
            backgroundColor: theme.isDark ? "#442222" : "#fde7e9", 
            color: theme.isDark ? "#ff9999" : "#d83b01", 
            borderRadius: "4px",
            fontSize: "12px",
            margin: "8px 0"
          }}>
            错误: {error}
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* 3. Footer: 固定在底部，不使用 absolute */}
      <div style={{ 
        padding: "8px", 
        backgroundColor: theme.isDark ? "#333333" : "#f3f2f1",
        borderTop: `1px solid ${theme.borderColor}`,
        flexShrink: 0,
        boxSizing: "border-box"
      }}>
        <div style={{ display: "flex", gap: "8px", width: "100%" }}>
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            onCompositionStart={() => { isComposing.current = true; }}
            onCompositionEnd={() => { isComposing.current = false; }}
            placeholder={isLoading ? "AI 正在响应中..." : "输入您的问题..."}
            disabled={isLoading}
            style={{ 
              flex: 1, 
              padding: "6px 10px", 
              borderRadius: "4px", 
              border: `1px solid ${theme.borderColor}`, 
              backgroundColor: isLoading ? (theme.isDark ? "#333333" : "#f5f5f5") : (theme.isDark ? "#444444" : "white"),
              color: theme.textColor,
              resize: "none",
              height: "36px",
              fontFamily: "inherit",
              fontSize: "13px",
              boxSizing: "border-box",
              opacity: isLoading ? 0.7 : 1
            }}
          />
          <button
            onClick={isLoading ? handleStop : () => handleSend()}
            disabled={!isLoading && !input.trim()}
            style={{
              padding: "0 12px",
              backgroundColor: !isLoading && !input.trim() ? (theme.isDark ? "#555555" : "#c8c6c4") : (isLoading ? "#d83b01" : theme.primaryColor),
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: !isLoading && !input.trim() ? "default" : "pointer",
              fontSize: "13px",
              minWidth: "60px"
            }}
          >
            {isLoading ? "停止" : "发送"}
          </button>
        </div>
      </div>
    </div>
  );
};

export default ChatPage;
