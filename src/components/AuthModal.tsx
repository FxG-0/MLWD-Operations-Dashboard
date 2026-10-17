import React, { useState } from "react";
import { motion, AnimatePresence } from "motion/react";
import { ShieldCheck, LogIn, X } from "lucide-react";
import { entraSignInRedirect } from "../lib/entraAuth";
import { UserAccount } from "../types";

interface AuthModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentUser: UserAccount;
  onSelectUser: (user: UserAccount) => void;
  isLightMode?: boolean;
}

export default function AuthModal({ isOpen, onClose, isLightMode = false }: AuthModalProps) {
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");

  if (!isOpen) return null;

  const handleSignIn = async () => {
    setErrorMsg("");
    setLoading(true);
    try {
      // Full-page redirect to Microsoft's login page. The page reloads;
      // App.tsx picks up the signed-in account on return via entraHandleReturn().
      await entraSignInRedirect();
    } catch (err: any) {
      console.error("Entra sign-in error:", err);
      setErrorMsg(err?.message || "Sign-in failed. Please try again.");
      setLoading(false);
    }
  };

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm px-4">
        <motion.div
          initial={{ opacity: 0, scale: 0.96, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.96, y: 10 }}
          className={`w-full max-w-sm rounded-2xl border p-6 relative ${
            isLightMode
              ? "bg-white border-slate-200 text-slate-800"
              : "bg-[#0d1017] border-slate-800 text-slate-200"
          }`}
        >
          <button
            onClick={onClose}
            className="absolute top-3 right-3 p-1.5 rounded-lg text-slate-400 hover:text-slate-200 hover:bg-white/5 cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>

          <div className="flex flex-col items-center text-center gap-3 mb-6">
            <div className="w-12 h-12 rounded-2xl bg-blue-600/10 border border-blue-600/30 flex items-center justify-center">
              <ShieldCheck className="w-6 h-6 text-blue-500" />
            </div>
            <div>
              <h2 className="text-base font-bold">MLWD Tools Dashboard</h2>
              <p className="text-xs text-slate-400 mt-1">
                Sign in with your Halliburton account to continue.
              </p>
            </div>
          </div>

          {errorMsg && (
            <div className="mb-4 text-xs text-red-400 bg-red-500/10 border border-red-500/20 rounded-xl px-3 py-2">
              {errorMsg}
            </div>
          )}

          <button
            onClick={handleSignIn}
            disabled={loading}
            className="w-full py-2.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold transition-all cursor-pointer shadow-lg shadow-blue-600/20 flex items-center justify-center gap-2 disabled:opacity-60"
          >
            <LogIn className="w-4 h-4" />
            <span>{loading ? "Redirecting to Microsoft..." : "Sign in with Microsoft"}</span>
          </button>

          <p className="text-[11px] text-slate-500 text-center mt-4">
            New here? Signing in automatically sends an access request — an
            admin will review it before you can view or edit data.
          </p>
        </motion.div>
      </div>
    </AnimatePresence>
  );
}
