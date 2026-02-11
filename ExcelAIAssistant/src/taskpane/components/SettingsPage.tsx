import * as React from "react";
import { ApiConfig, ApiType } from "./types";

interface SettingsPageProps {
  configs: ApiConfig[];
  activeConfigId: string | null;
  onSave: (configs: ApiConfig[], activeId: string | null) => void;
  theme: {
    isDark: boolean;
    primaryColor: string;
    bgColor: string;
    textColor: string;
    borderColor: string;
  };
}

const SettingsPage: React.FC<SettingsPageProps> = ({ configs, activeConfigId, onSave, theme }) => {
  const [localConfigs, setLocalConfigs] = React.useState<ApiConfig[]>(configs);

  const addConfig = () => {
    const newConfig: ApiConfig = {
      id: Date.now().toString(),
      name: "新配置",
      type: "openai",
      apiKey: "",
      endpoint: "https://api.openai.com/v1",
      model: "gpt-4o",
    };
    const updated = [...localConfigs, newConfig];
    setLocalConfigs(updated);
    onSave(updated, activeConfigId || newConfig.id);
  };

  const updateConfig = (id: string, updates: Partial<ApiConfig>) => {
    const updated = localConfigs.map((c) => (c.id === id ? { ...c, ...updates } : c));
    setLocalConfigs(updated);
    onSave(updated, activeConfigId);
  };

  const deleteConfig = (id: string) => {
    const updated = localConfigs.filter((c) => c.id !== id);
    const newActiveId = activeConfigId === id ? (updated.length > 0 ? updated[0].id : null) : activeConfigId;
    setLocalConfigs(updated);
    onSave(updated, newActiveId);
  };

  const selectActive = (id: string) => {
    onSave(localConfigs, id);
  };

  const inputStyle: React.CSSProperties = {
    width: "100%",
    padding: "6px",
    border: `1px solid ${theme.borderColor}`,
    borderRadius: "4px",
    boxSizing: "border-box",
    fontSize: "13px",
    backgroundColor: theme.isDark ? "#3c3c3c" : "white",
    color: theme.textColor,
    fontFamily: "inherit"
  };

  const labelStyle: React.CSSProperties = {
    display: "block",
    fontSize: "11px",
    color: theme.isDark ? "#bbbbbb" : "#605e5c",
    marginBottom: "2px"
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "15px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h2 style={{ fontSize: "16px", margin: 0, color: theme.textColor }}>API 配置</h2>
        <button 
          onClick={addConfig}
          style={{
            padding: "5px 10px",
            backgroundColor: theme.primaryColor,
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: "pointer",
            fontSize: "13px"
          }}
        >
          添加配置
        </button>
      </div>

      {localConfigs.map((config) => (
        <div 
          key={config.id} 
          style={{ 
            padding: "12px", 
            backgroundColor: theme.isDark ? "#333333" : "white", 
            borderRadius: "6px", 
            boxShadow: "0 1px 3px rgba(0,0,0,0.2)",
            border: activeConfigId === config.id ? `2px solid ${theme.primaryColor}` : `1px solid ${theme.borderColor}`,
            position: "relative"
          }}
        >
          <div style={{ position: "absolute", top: "8px", right: "8px" }}>
            <label style={{ fontSize: "11px", cursor: "pointer", display: "flex", alignItems: "center", gap: "4px", color: theme.textColor }}>
              <input 
                type="radio" 
                checked={activeConfigId === config.id} 
                onChange={() => selectActive(config.id)}
              />
              当前使用
            </label>
          </div>

          <div style={{ marginBottom: "8px", marginTop: "5px" }}>
            <label style={labelStyle}>配置名称</label>
            <input
              value={config.name}
              onChange={(e) => updateConfig(config.id, { name: e.target.value })}
              style={inputStyle}
            />
          </div>

          <div style={{ marginBottom: "8px" }}>
            <label style={labelStyle}>服务类型</label>
            <select
              value={config.type}
              onChange={(e) => updateConfig(config.id, { type: e.target.value as ApiType })}
              style={inputStyle}
            >
              <option value="openai">OpenAI 兼容 (DeepSeek/Local)</option>
              <option value="gemini">Google Gemini</option>
              <option value="doubao">火山引擎 (豆包)</option>
            </select>
          </div>

          <div style={{ marginBottom: "8px" }}>
            <label style={labelStyle}>API Key</label>
            <input
              type="password"
              value={config.apiKey}
              onChange={(e) => updateConfig(config.id, { apiKey: e.target.value })}
              style={inputStyle}
            />
          </div>

          {config.type !== "gemini" && (
            <div style={{ marginBottom: "8px" }}>
              <label style={labelStyle}>接口地址 (Endpoint)</label>
              <input
                value={config.endpoint}
                onChange={(e) => updateConfig(config.id, { endpoint: e.target.value })}
                placeholder={config.type === "doubao" ? "https://ark.cn-beijing.volces.com/api/v3" : "https://api.openai.com/v1"}
                style={inputStyle}
              />
            </div>
          )}

          <div style={{ marginBottom: "12px" }}>
            <label style={labelStyle}>
              {config.type === "doubao" ? "推理端点 ID (Endpoint ID)" : "模型名称 (Model)"}
            </label>
            <input
              value={config.model}
              onChange={(e) => updateConfig(config.id, { model: e.target.value })}
              placeholder={config.type === "doubao" ? "ep-202xxxx-xxxx" : "gpt-4o / deepseek-chat"}
              style={inputStyle}
            />
          </div>

          <button 
            onClick={() => deleteConfig(config.id)}
            style={{
              width: "100%",
              padding: "6px",
              backgroundColor: "transparent",
              color: theme.isDark ? "#ff9999" : "#d83b01",
              border: `1px solid ${theme.isDark ? "#ff9999" : "#d83b01"}`,
              borderRadius: "4px",
              cursor: "pointer",
              fontSize: "12px"
            }}
          >
            删除配置
          </button>
        </div>
      ))}
    </div>
  );
};

export default SettingsPage;
