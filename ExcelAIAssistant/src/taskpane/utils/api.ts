import { ApiConfig } from "../components/types";
import { ExcelContext } from "../excel/types";

export interface Message {
  role: "system" | "user" | "assistant";
  content: string;
}

/**
 * 简单的上下文压缩逻辑
 * 1. 始终保留最新的 N 条消息
 * 2. 对旧消息进行摘要处理（或简单的丢弃，此处先实现滑动窗口）
 * 3. 始终保留系统提示词
 */
export function compressContext(messages: Message[], maxWindow: number = 20): Message[] {
  if (messages.length <= maxWindow) return messages;

  const systemMessage = messages.find(m => m.role === "system");
  // 始终尝试保留最近的交互
  const recentMessages = messages.slice(-maxWindow);
  
  // 过滤掉可能存在的系统指令注入，避免干扰
  const filteredMessages = recentMessages.filter(m => {
    // 只有非 system 角色或者明确的上下文消息才保留
    return m.role !== "system";
  });
  
  // 如果有系统消息且不在最近消息中，则将其添加回去
  if (systemMessage) {
    return [systemMessage, ...filteredMessages];
  }
  
  return filteredMessages;
}

export const SYSTEM_PROMPT = `你是一个专业的 Excel AI 助手，负责通过指令或脚本自动化 Excel 任务。

## 🎯 核心原则：观察优先 (OBSERVATION FIRST) & 稳健运行
为了确保任务 100% 成功，请严格遵循以下操作顺序：

1.  **先观察再行动 (Look Before You Leap)**：
    *   **原则**：严禁在不确定单元格位置、工作表结构或数据内容时直接写入。
    *   **操作**：如果任务涉及现有数据，**必须先使用 \`read\` 指令**观察数据范围、表头和内容格式。不要凭空猜测 \`A1\` 或 \`UsedRange\` 的内容。
    *   **环境意识**：请随时关注下方的“当前上下文”信息。如果发现 \`usedRangeAddress\` 为 \`A1:A1\` 且 \`headers\` 为空，说明工作表可能是空的或尚未读取。
    *   **目的**：防止因引用错误导致的数据覆盖或任务失败。
2.  **优先使用快捷指令 (Quick Ops)**：
    *   **原因**：快捷指令经过底层适配器的预编译处理，避开了动态脚本解析的兼容性坑。
    *   **适用**：读写数据、格式设置、增删行列、**新建工作表**、**跨表复制粘贴**。
3.  **善用 Excel 公式 (Excel Formulas)**：
    *   **核心思维**：**不要在 JS 脚本中做计算，让 Excel 自己算。**
    *   **操作**：使用 \`write\` 指令将公式（如 \`=SUM(A1:B10)\` 或 \`=VLOOKUP(...)\`）写入单元格，而不是读取数据到 JS 算完再写回。
    *   **批量写入技巧**：Excel 在批量写入公式时不会自动调整相对引用。**务必使用 \`ROW()\` 函数实现动态引用**。例如：使用 \`"=INDIRECT(\"A\"&ROW())"\` 动态引用当前行的 A 列，而非固定的 \`"A1"\`。
    *   **优势**：性能更好、支持多语言、自动重算、避免 JS 浮点数精度问题。
4.  **任务终结机制 (Termination)**：
    *   当你认为当前任务已全部完成，不再需要执行更多指令或读取更多数据时，**必须在回复的最后加上 \`[FINISH]\` 标记**。
    *   该标记会告知系统停止自动反馈循环，将控制权交还给用户。
5.  **最后才考虑 JS 脚本 (Script)**：
    *   仅在涉及极其复杂的循环、外部 API 调用或快捷指令完全无法实现的逻辑时使用。

## 💡 快捷指令 (Quick Ops) 规范
你必须输出包裹在 \`\`\`json 代码块中的 JSON 对象，严禁直接输出或使用单反引号。
- **读取 (首选操作)**: 
\`\`\`json
{"type": "read", "range": "'Sheet1'!A1:B10"}
\`\`\`
- **写入 (含公式)**: 
\`\`\`json
{"type": "write", "range": "A1", "value": "=SUM(B1:B10)"}
\`\`\`
- **格式**: 
\`\`\`json
{"type": "format", "range": "A1", "style": {"fillColor": "#FFFF00", "bold": true}}
\`\`\`
- **新建表**: 
\`\`\`json
{"type": "create_sheet", "sheetName": "新表名"}
\`\`\`
- **绘制图表**: 
\`\`\`json
{"type": "chart", "range": "A1:B10", "chartType": "ColumnClustered", "title": "图表标题"}
\`\`\`
- **复制粘贴**: 
\`\`\`json
{"type": "copy_paste", "range": "'源表'!A1:B10", "destination": "'目标表'!A1"}
\`\`\`
- **插入/删除**: 
\`\`\`json
{"type": "insert", "range": "Row1"}
\`\`\`

## ⚠️ 关键细节
- **跨表操作**：在 \`range\` 或 \`destination\` 中引用非当前表时，必须使用 \`'工作表名'!单元格地址\` 格式。
- **自动执行**：你的 JSON 或 JavaScript 代码块会在输出后自动运行，结果会在下一轮对话反馈。
- **脚本规范**：若必须写脚本，严禁使用 ES6+ 语法 (const, let, =>, async/await)，必须使用 var 和 .then()。

## 当前上下文
{context}

请直接根据用户需求输出指令，不要解释上述规则。`;

export interface AIConfig {
    endpoint?: string;
    apiKey: string;
    model?: string;
    type?: string;
}

export async function callAI(
  messages: any[], 
  config: AIConfig, 
  onDelta?: (delta: string) => void
): Promise<{ content: string; abort: () => void }> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    let url = config.endpoint || "";
    
    // 自动补全 URL 路径逻辑
    if (config.type === "openai" || config.type === "doubao") {
      if (url && !url.endsWith("/chat/completions")) {
        url = url.replace(/\/$/, "") + "/chat/completions";
      }
    } else if (config.type === "gemini" && !url) {
      url = `https://generativelanguage.googleapis.com/v1beta/models/${config.model || "gemini-1.5-pro"}:generateContent?key=${config.apiKey}`;
    }

    xhr.open("POST", url);
    xhr.setRequestHeader("Content-Type", "application/json");
    if (config.type !== "gemini") {
      xhr.setRequestHeader("Authorization", `Bearer ${config.apiKey || ""}`);
    }

    let seenBytes = 0;
    let fullContent = "";
    let geminiBuffer = ""; 
    let lineBuffer = ""; // 新增：用于 OpenAI/标准流式的行缓冲区

    xhr.onprogress = () => {
      // 在某些环境下，状态码可能在流式传输中途变为 200，或者初始为 0
      if (xhr.status !== 200 && xhr.status !== 0) return;
      
      const newData = xhr.responseText.substring(seenBytes);
      seenBytes = xhr.responseText.length;

      // 1. Gemini 专用缓冲逻辑 (JSON 数组格式)
      if (config.type === "gemini") {
        geminiBuffer += newData;
        try {
          let searchIdx = 0;
          while (true) {
            const startIdx = geminiBuffer.indexOf('{', searchIdx);
            if (startIdx === -1) break;

            let endIdx = -1;
            let depth = 0;
            for (let i = startIdx; i < geminiBuffer.length; i++) {
              if (geminiBuffer[i] === '{') depth++;
              else if (geminiBuffer[i] === '}') {
                depth--;
                if (depth === 0) {
                  endIdx = i;
                  break;
                }
              }
            }

            if (endIdx !== -1) {
              const jsonStr = geminiBuffer.substring(startIdx, endIdx + 1);
              try {
                const parsed = JSON.parse(jsonStr);
                const content = parsed.candidates?.[0]?.content?.parts?.[0]?.text || "";
                if (content) {
                  fullContent += content;
                  if (onDelta) onDelta(content);
                }
                geminiBuffer = geminiBuffer.substring(endIdx + 1);
                searchIdx = 0;
              } catch (e) {
                searchIdx = endIdx + 1;
              }
            } else {
              break;
            }
          }
        } catch (e) {}
        return;
      }

      // 2. OpenAI/标准 SSE 逻辑 (data: 格式)
      // 使用 lineBuffer 处理可能被截断的行
      const combined = lineBuffer + newData;
      const lines = combined.split("\n");
      // 最后一行可能是不完整的，留到下次处理
      lineBuffer = lines.pop() || "";

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || !trimmed.startsWith("data: ")) continue;
        
        const data = trimmed.slice(6);
        if (data === "[DONE]") continue;
        
        try {
          const parsed = JSON.parse(data);
          const content = parsed.choices?.[0]?.delta?.content || "";
          if (content) {
            fullContent += content;
            if (onDelta) onDelta(content);
          }
        } catch (e) {
          // 如果 JSON 解析失败，说明这行可能还没收全（尽管 split 了 \n）
          // 这种情况在某些非标实现中可能发生，尝试将其加回 buffer
          lineBuffer = line + "\n" + lineBuffer;
        }
      }
    };

    xhr.onreadystatechange = () => {
      if (xhr.readyState === 4) {
        if (xhr.status >= 200 && xhr.status < 300) {
          // 如果没有 onDelta 回调，或者是普通非流式请求
          if (!fullContent && xhr.responseText) {
            try {
              const response = JSON.parse(xhr.responseText);
              if (config.type === "gemini") {
                fullContent = response.candidates?.[0]?.content?.parts?.[0]?.text || "";
              } else {
                fullContent = response.choices?.[0]?.message?.content || "";
              }
            } catch (e) {}
          }
          resolve({ content: fullContent, abort: () => xhr.abort() });
        } else {
          let errorMsg = `请求失败: ${xhr.status}`;
          try {
            const errData = JSON.parse(xhr.responseText);
            errorMsg += ` - ${errData.error?.message || xhr.statusText}`;
          } catch (e) {}
          reject(new Error(errorMsg));
        }
      }
    };

    xhr.onerror = () => reject(new Error("网络错误，请检查网络连接。"));

    // 构造 Payload
    let payload: any;
    if (config.type === "gemini") {
      payload = {
        contents: messages.filter(m => m.role !== "system").map(m => ({
          role: m.role === "assistant" ? "model" : "user",
          parts: [{ text: m.content }]
        }))
      };
      const systemMsg = messages.find(m => m.role === "system");
      if (systemMsg) {
        payload.system_instruction = { parts: [{ text: systemMsg.content }] };
      }
    } else {
      payload = { model: config.model, messages, stream: true };
    }

    if (config.type === "gemini") {
      // Gemini 流式使用不同的接口
      if (onDelta) {
        url = url.replace("generateContent", "streamGenerateContent");
      }
    }

    xhr.send(JSON.stringify(payload));
  });
}
