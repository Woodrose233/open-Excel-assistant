import * as React from "react";

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: "20px", color: "#d83b01", fontFamily: "sans-serif" }}>
          <h2>出现了一些问题</h2>
          <p>很抱歉，插件在加载时遇到了错误。</p>
          <pre style={{ whiteSpace: "pre-wrap", background: "#f3f2f1", padding: "10px" }}>
            {this.state.error?.toString()}
          </pre>
          <p>建议尝试：</p>
          <ul>
            <li>确保已安装最新的 Microsoft Edge WebView2 运行时</li>
            <li>刷新页面或重启 Excel</li>
          </ul>
        </div>
      );
    }

    return this.props.children;
  }
}
