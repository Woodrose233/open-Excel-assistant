export type ApiType = "openai" | "gemini" | "doubao";

export interface ApiConfig {
  id: string;
  name: string;
  type: ApiType;
  apiKey: string;
  endpoint?: string; // 针对 OpenAI 兼容地址或豆包特定地址
  model?: string;     // 模型名称 (如 gpt-4o, deepseek-chat, gemini-1.5-pro 等)
}

export interface AppSettings {
  activeConfigId: string | null;
  configs: ApiConfig[];
}

export const DEFAULT_SETTINGS: AppSettings = {
  activeConfigId: null,
  configs: [],
};
