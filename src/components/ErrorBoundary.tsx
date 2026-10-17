import React from "react";
import { AlertOctagon, RotateCcw } from "lucide-react";

interface Props {
  children: React.ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export default class ErrorBoundary extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = {
      hasError: false,
      error: null
    };
  }

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("Uncaught enterprise applet error:", error, errorInfo);
  }

  private handleReset = () => {
    localStorage.clear();
    window.location.reload();
  };

  public override render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-[#080B11] text-slate-100 flex flex-col items-center justify-center p-6 font-sans antialiased">
          <div className="max-w-md w-full bg-[#0F1420] border border-red-900/40 rounded-2xl p-6 shadow-xl space-y-4 text-center">
            <div className="mx-auto w-12 h-12 rounded-full bg-red-950/40 text-red-400 border border-red-900/30 flex items-center justify-center">
              <AlertOctagon className="w-6 h-6" />
            </div>

            <div className="space-y-1">
              <h2 className="text-sm font-bold uppercase tracking-wider text-slate-300 font-display">
                Application Exception Encountered
              </h2>
              <p className="text-xs text-slate-400">
                An unexpected front-end render state occurred in the control panel.
              </p>
            </div>

            {this.state.error && (
              <div className="p-3 bg-red-950/20 text-red-300 text-[10px] font-mono rounded-lg border border-red-900/20 text-left overflow-x-auto max-h-[140px]">
                {this.state.error.stack || this.state.error.message}
              </div>
            )}

            <div className="flex flex-col gap-2 pt-2">
              <button
                type="button"
                onClick={() => this.setState({ hasError: false, error: null })}
                className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-bold bg-blue-600 hover:bg-blue-500 text-white rounded-xl cursor-pointer transition-colors shadow-sm"
              >
                <span>Attempt Hot Recovery</span>
              </button>

              <button
                type="button"
                onClick={this.handleReset}
                className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl cursor-pointer transition-colors border border-slate-700/50"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Reset Cache & Refresh</span>
              </button>
            </div>
            
            <p className="text-[10px] text-slate-500 font-mono">
              Your uploaded data will be safely restored to system defaults upon cache reset.
            </p>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
