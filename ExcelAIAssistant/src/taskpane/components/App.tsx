import * as React from "react";
import { AppSettings, DEFAULT_SETTINGS, ApiConfig } from "./types";
import { loadSettings, saveSettings } from "./storage";
import SettingsPage from "./SettingsPage";
import ChatPage from "./ChatPage";
import ScriptTestPage from "./ScriptTestPage";
import DebugTestPage from "./DebugTestPage";
import StreamDebugPage from "./StreamDebugPage";

interface AppProps {
  title: string;
}

const App: React.FC<AppProps> = () => {
  const [settings, setSettings] = React.useState<AppSettings>(DEFAULT_SETTINGS);
  const [activeTab, setActiveTab] = React.useState<"chat" | "settings" | "test" | "debug" | "stream-debug">("chat");
  const [isLoaded, setIsLoaded] = React.useState(false);
  const [theme, setTheme] = React.useState<Office.OfficeTheme | null>(null);

  React.useEffect(() => {
    const initialize = async () => {
      try {
        // 获取 Office 主题，增加防御性代码
        if (typeof Office !== "undefined" && Office.context && Office.context.officeTheme) {
          setTheme(Office.context.officeTheme);
        }

        const savedSettings = loadSettings();
        setSettings(savedSettings);
        
        // 如果没有配置，强制进入设置页面
        if (savedSettings.configs.length === 0) {
          setActiveTab("settings");
        } else {
          setActiveTab("chat");
        }
      } catch (err) {
        console.error("Failed to initialize app", err);
      } finally {
        // 无论如何，在 2 秒后强制结束加载状态，防止页面死锁
        setTimeout(() => setIsLoaded(true), 100);
      }
    };

    initialize();
  }, []);

  const handleSaveConfigs = (configs: ApiConfig[], activeConfigId: string | null) => {
    const newSettings = { ...settings, configs, activeConfigId };
    setSettings(newSettings);
    saveSettings(newSettings);
  };

  const activeConfig = settings.configs.find((c) => c.id === settings.activeConfigId) || (settings.configs.length > 0 ? settings.configs[0] : null);

  // 计算颜色
  const isDark = theme && theme.bodyBackgroundColor ? (
    // 简单的亮度检测
    parseInt(theme.bodyBackgroundColor.substring(1, 3), 16) < 128
  ) : false;

  const primaryColor = "#217346"; // Excel 标志绿
  const bgColor = theme?.bodyBackgroundColor || (isDark ? "#2b2b2b" : "#f3f2f1");
  const textColor = isDark ? "#ffffff" : (theme?.bodyForegroundColor || "#323130");
  const navBgColor = isDark ? "#3c3c3c" : (theme?.controlBackgroundColor || "#ffffff");
  const borderColor = isDark ? "#555555" : "#edebe9";

  const subTabStyle: React.CSSProperties = {
    flex: 1,
    padding: "6px",
    border: "none",
    backgroundColor: "transparent",
    color: isDark ? "#e0e0e0" : "#323130",
    fontSize: "12px",
    cursor: "pointer"
  };

  if (!isLoaded) {
    return <div style={{ padding: "20px", color: textColor, backgroundColor: bgColor }}>加载中...</div>;
  }

  return (
    <div style={{ 
      display: "flex", 
      flexDirection: "column", 
      height: "100vh", 
      fontFamily: "'Segoe UI', 'Microsoft YaHei', sans-serif",
      backgroundColor: bgColor,
      color: textColor
    }}>
      <nav style={{ 
        display: "flex", 
        backgroundColor: navBgColor, 
        borderBottom: `1px solid ${borderColor}` 
      }}>
        <button 
          onClick={() => setActiveTab("chat")}
          style={{
            flex: 1,
            padding: "8px",
            border: "none",
            backgroundColor: "transparent",
            borderBottom: activeTab === "chat" ? `2px solid ${primaryColor}` : "none",
            color: activeTab === "chat" ? (isDark ? "#40bc74" : primaryColor) : (isDark ? "#e0e0e0" : "#323130"),
            fontWeight: activeTab === "chat" ? "600" : "400",
            cursor: "pointer",
            fontSize: "13px"
          }}
        >
          聊天
        </button>
        <button 
          onClick={() => setActiveTab("settings")}
          style={{
            flex: 1,
            padding: "8px",
            border: "none",
            backgroundColor: "transparent",
            borderBottom: activeTab === "settings" ? `2px solid ${primaryColor}` : "none",
            color: activeTab === "settings" ? (isDark ? "#40bc74" : primaryColor) : (isDark ? "#e0e0e0" : "#323130"),
            fontWeight: activeTab === "settings" ? "600" : "400",
            cursor: "pointer",
            fontSize: "13px"
          }}
        >
          设置
        </button>
        <button 
          onClick={() => setActiveTab("debug")}
          style={{
            flex: 1,
            padding: "8px",
            border: "none",
            backgroundColor: "transparent",
            borderBottom: activeTab === "debug" ? `2px solid ${primaryColor}` : "none",
            color: activeTab === "debug" ? (isDark ? "#40bc74" : primaryColor) : (isDark ? "#e0e0e0" : "#323130"),
            fontWeight: activeTab === "debug" ? "600" : "400",
            cursor: "pointer",
            fontSize: "13px"
          }}
        >
          调试
        </button>
      </nav>

      <main style={{ flex: 1, position: "relative" }}>
        {activeTab === "chat" && (
          <ChatPage config={activeConfig} theme={{ isDark, primaryColor, bgColor, textColor, borderColor }} />
        )}
        {activeTab === "settings" && (
          <SettingsPage 
            configs={settings.configs} 
            activeConfigId={settings.activeConfigId}
            onSave={handleSaveConfigs} 
            theme={{ isDark, primaryColor, bgColor, textColor, borderColor }}
          />
        )}
        {activeTab === "debug" && (
          <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
            <div style={{ display: "flex", borderBottom: `1px solid ${borderColor}`, backgroundColor: navBgColor }}>
              <button 
                onClick={() => setActiveTab("debug")}
                style={{ ...subTabStyle, borderBottom: activeTab === "debug" ? `2px solid ${primaryColor}` : "none" }}
              >
                功能调试
              </button>
              <button 
                onClick={() => setActiveTab("stream-debug")}
                style={{ ...subTabStyle, borderBottom: activeTab === "stream-debug" ? `2px solid ${primaryColor}` : "none" }}
              >
                流式调试
              </button>
            </div>
            <DebugTestPage />
          </div>
        )}
        {activeTab === "stream-debug" && (
          <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
            <div style={{ display: "flex", borderBottom: `1px solid ${borderColor}`, backgroundColor: navBgColor }}>
              <button 
                onClick={() => setActiveTab("debug")}
                style={{ ...subTabStyle, borderBottom: activeTab === "debug" ? `2px solid ${primaryColor}` : "none" }}
              >
                功能调试
              </button>
              <button 
                onClick={() => setActiveTab("stream-debug")}
                style={{ ...subTabStyle, borderBottom: activeTab === "stream-debug" ? `2px solid ${primaryColor}` : "none" }}
              >
                流式调试
              </button>
            </div>
            <StreamDebugPage config={activeConfig} />
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
