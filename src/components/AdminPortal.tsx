import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import {
  ShieldCheck,
  Eye,
  UserPlus,
  CheckCircle2,
  Trash2,
  Lock,
  Mail,
  User,
  Search,
  RefreshCw,
  AlertCircle,
  Shield,
  Edit3,
  X,
  Key,
  Check,
  UserCheck,
  ChevronRight,
  Sparkles,
  Users,
  Radio,
  UserX,
  Clock
} from "lucide-react";
import { authFetch } from "../lib/entraAuth";
import { UserAccount, UserRole, UserStatus, UserPermissions } from "../types";

interface AdminPortalProps {
  isOpen: boolean;
  onClose: () => void;
  currentUser: UserAccount;
  isLightMode?: boolean;
  onUserListChanged?: () => void;
  activeUserCount?: number;
  activeUsersList?: any[];
}

export default function AdminPortal({
  isOpen,
  onClose,
  currentUser,
  isLightMode = false,
  onUserListChanged,
  activeUserCount = 1,
  activeUsersList = []
}: AdminPortalProps) {
  const [users, setUsers] = useState<UserAccount[]>([]);
  const [loading, setLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [toastMsg, setToastMsg] = useState("");
  const [errorMsg, setErrorMsg] = useState("");

  // New User Form state
  const [isAddingUser, setIsAddingUser] = useState(false);
  const [newName, setNewName] = useState("");
  const [newUsername, setNewUsername] = useState("");
  const [newEmail, setNewEmail] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [newRole, setNewRole] = useState<UserRole>("viewer");
  const [newPermissions, setNewPermissions] = useState<UserPermissions>("read");

  // Editing existing user modal state
  const [editingUser, setEditingUser] = useState<UserAccount | null>(null);
  const [editName, setEditName] = useState("");
  const [editUsername, setEditUsername] = useState("");
  const [editEmail, setEditEmail] = useState("");
  const [editPassword, setEditPassword] = useState("");
  const [editRole, setEditRole] = useState<UserRole>("viewer");
  const [editPermissions, setEditPermissions] = useState<UserPermissions>("read");
  const [editStatus, setEditStatus] = useState<UserStatus>("approved");

  // Status filter state: all, pending, approved
  const [statusFilter, setStatusFilter] = useState<"all" | "pending" | "approved">("all");

  const fetchUsers = React.useCallback(() => {
    return authFetch("/api/users")
      .then((res) => res.json())
      .then((data) => {
        if (Array.isArray(data?.users)) {
          setUsers(data.users);
        }
      })
      .catch(() => {});
  }, []);

  useEffect(() => {
    if (!isOpen) return;
    setLoading(true);
    fetchUsers().finally(() => setLoading(false));

    // Poll while the panel's open so approvals/role changes from elsewhere show up
    const interval = setInterval(fetchUsers, 5000);
    return () => clearInterval(interval);
  }, [isOpen, fetchUsers]);

  if (!isOpen) return null;

  const showToast = (msg: string) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(""), 3500);
  };

  const canEditUser = currentUser?.role === "admin" || currentUser?.role === "editor" || currentUser?.permissions === "both" || currentUser?.permissions === "write";

  const handleUpdateUserQuick = async (
    targetUid: string,
    updates: Partial<UserAccount>
  ) => {
    if (!canEditUser) return;
    const existing = users.find((u) => u.id === targetUid);
    try {
      const res = await authFetch(`/api/users/row`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ user: { ...(existing || {}), id: targetUid, ...updates } })
      });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || "Failed to update user account.");
      }
      setUsers((prev) =>
        prev.map((u) => (u.id === targetUid ? { ...u, ...updates } : u))
      );
      if (updates.status === "approved") {
        showToast("User access approved!");
      } else if (updates.status === "pending") {
        showToast("User access revoked (set to pending).");
      } else if (updates.status === "rejected") {
        showToast("User access declined.");
      } else {
        showToast("Account updated successfully!");
      }
      onUserListChanged?.();
    } catch (e: any) {
      console.error("Error updating user:", e);
      setErrorMsg(e.message || "Failed to update user account.");
    }
  };

  const handleDeleteUser = async (targetUid: string, name: string, email: string) => {
    if (!canEditUser) return;
    if (targetUid === currentUser.id) {
      setErrorMsg("Cannot delete your own active account.");
      return;
    }
    if (!confirm(`Are you sure you want to remove account '${name}'?`)) return;

    try {
      const res = await authFetch(`/api/users/row/${targetUid}`, { method: "DELETE" });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || "Failed to delete account.");
      }
      setUsers((prev) => prev.filter((u) => u.id !== targetUid));
      showToast(`Account '${name}' deleted.`);
      onUserListChanged?.();
    } catch (e: any) {
      console.error("Error deleting user:", e);
      setErrorMsg(e.message || "Network error deleting account.");
    }
  };

  const openEditModal = (u: UserAccount) => {
    if (!canEditUser) return;
    setEditingUser(u);
    setEditName(u.name || "");
    setEditUsername(u.username || u.email.split("@")[0] || "");
    setEditEmail(u.email || "");
    setEditPassword("");
    setEditRole(u.role || "viewer");
    setEditPermissions(u.permissions || "read");
    setEditStatus(u.status || "approved");
  };

  const handleSaveEditUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingUser || !editingUser.id) return;

    const updates = {
      name: editName.trim(),
      username: editUsername.trim(),
      email: editEmail.trim(),
      role: editRole,
      permissions: editPermissions,
      status: editStatus
    };

    try {
      const res = await authFetch(`/api/users/row`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ user: { ...editingUser, ...updates } })
      });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || "Failed to save account edits.");
      }
      setUsers((prev) => prev.map((u) => (u.id === editingUser.id ? { ...u, ...updates } : u)));
      showToast(`Account '${editName.trim()}' updated!`);
      setEditingUser(null);
      onUserListChanged?.();
    } catch (e: any) {
      console.error("Error updating user:", e);
      setErrorMsg(e.message || "Error saving account edits.");
    }
  };

  const filteredUsers = users.filter((u) => {
    if (statusFilter === "pending" && u.status !== "pending") return false;
    if (statusFilter === "approved" && u.status !== "approved") return false;
    const q = searchQuery.toLowerCase();
    if (!q) return true;
    return (
      (u.name && u.name.toLowerCase().includes(q)) ||
      (u.email && u.email.toLowerCase().includes(q)) ||
      (u.username && u.username.toLowerCase().includes(q)) ||
      (u.role && u.role.toLowerCase().includes(q)) ||
      (u.status && u.status.toLowerCase().includes(q))
    );
  });

  const adminCount = users.filter((u) => u.role === "admin").length;
  const editorCount = users.filter((u) => u.role === "editor").length;
  const userCount = users.filter((u) => u.role === "viewer").length;
  const pendingCount = users.filter((u) => u.status === "pending").length;
  const approvedCount = users.filter((u) => u.status === "approved").length;

  return (
    <AnimatePresence>
      <div
        className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center z-[110] p-4"
        onClick={onClose}
      >
        <motion.div
          initial={{ opacity: 0, scale: 0.96, y: 15 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.96, y: 15 }}
          className="relative max-w-4xl w-full max-h-[90vh] rounded-2xl border border-slate-700/80 shadow-2xl overflow-hidden font-sans flex flex-col bg-[#121620] text-slate-100"
          onClick={(e) => e.stopPropagation()}
        >
          {/* Header */}
          <div className="flex items-center justify-between px-6 py-4 border-b border-slate-800/80 shrink-0 bg-[#141A29]">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-purple-600/15 text-purple-400 border border-purple-500/30">
                <ShieldCheck className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold tracking-tight font-display text-slate-100">
                  Admin Panel & User Account Manager
                </h3>
                <p className="text-xs text-slate-400">
                  Manage Firestore user roles, approve pending registrations, and create new accounts.
                </p>
              </div>
            </div>

            <button
              onClick={onClose}
              className="p-2 rounded-lg transition-colors cursor-pointer bg-slate-800/50 hover:bg-slate-800 text-slate-400 hover:text-white"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          {/* Toast / Error Banner */}
          {toastMsg && (
            <div className="px-6 py-2 bg-emerald-600/20 border-b border-emerald-500/30 text-emerald-300 text-xs font-semibold flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
              <span>{toastMsg}</span>
            </div>
          )}

          {errorMsg && (
            <div className="px-6 py-2 bg-rose-600/20 border-b border-rose-500/30 text-rose-300 text-xs font-semibold flex items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
                <span>{errorMsg}</span>
              </div>
              <button onClick={() => setErrorMsg("")} className="hover:text-white cursor-pointer">
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          )}

          {/* Stats Bar */}
          <div className="px-6 py-3 border-b border-slate-800/60 grid grid-cols-2 sm:grid-cols-5 gap-3 text-xs shrink-0 font-mono bg-slate-900/40">
            <div className="p-2.5 rounded-xl border border-emerald-500/30 flex items-center justify-between bg-emerald-950/20">
              <span className="text-emerald-300 flex items-center gap-1.5 font-bold">
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse shrink-0" />
                Active Online
              </span>
              <span className="font-bold text-sm text-emerald-400 font-mono">{activeUserCount}</span>
            </div>
            <div className="p-2.5 rounded-xl border border-slate-800 flex items-center justify-between bg-slate-900/80">
              <span className="text-slate-400">Total Accounts</span>
              <span className="font-bold text-sm text-white">{users.length}</span>
            </div>
            <div className="p-2.5 rounded-xl border border-purple-500/30 flex items-center justify-between bg-purple-950/20">
              <span className="text-purple-300">Admin</span>
              <span className="font-bold text-sm text-purple-400">{adminCount}</span>
            </div>
            <div className="p-2.5 rounded-xl border border-blue-500/30 flex items-center justify-between bg-blue-950/20">
              <span className="text-blue-300">Editor</span>
              <span className="font-bold text-sm text-blue-400">{editorCount}</span>
            </div>
            <div className="p-2.5 rounded-xl border border-amber-500/30 flex items-center justify-between bg-amber-950/20">
              <span className="text-amber-300">Pending</span>
              <span className="font-bold text-sm text-amber-400">{pendingCount}</span>
            </div>
          </div>

          {/* Controls Bar */}
          <div className="px-6 py-3 border-b border-slate-800/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shrink-0 bg-[#10141D]">
            <div className="relative flex-1 max-w-md">
              <Search className="w-3.5 h-3.5 absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search user by account name, email or status..."
                className="w-full h-8 pl-8 pr-3 text-xs rounded-xl border border-slate-700/80 bg-[#141824] text-slate-200 placeholder:text-slate-500 focus:border-blue-500 focus:outline-none transition-colors"
              />
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              {/* Status Filter Chips */}
              <div className="flex items-center gap-1 bg-[#141824] p-0.5 rounded-xl border border-slate-800 text-[11px] font-mono">
                <button
                  type="button"
                  onClick={() => setStatusFilter("all")}
                  className={`px-2.5 py-1 rounded-lg font-bold transition-colors cursor-pointer ${
                    statusFilter === "all"
                      ? "bg-slate-700 text-white"
                      : "text-slate-400 hover:text-slate-200"
                  }`}
                >
                  All ({users.length})
                </button>
                <button
                  type="button"
                  onClick={() => setStatusFilter("pending")}
                  className={`px-2.5 py-1 rounded-lg font-bold transition-colors cursor-pointer flex items-center gap-1.5 ${
                    statusFilter === "pending"
                      ? "bg-amber-600 text-white shadow"
                      : "text-amber-400/90 hover:text-amber-300"
                  }`}
                >
                  {pendingCount > 0 && <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-ping" />}
                  <span>Pending ({pendingCount})</span>
                </button>
                <button
                  type="button"
                  onClick={() => setStatusFilter("approved")}
                  className={`px-2.5 py-1 rounded-lg font-bold transition-colors cursor-pointer ${
                    statusFilter === "approved"
                      ? "bg-emerald-600 text-white shadow"
                      : "text-emerald-400/90 hover:text-emerald-300"
                  }`}
                >
                  Approved ({approvedCount})
                </button>
              </div>

              <button
                type="button"
                onClick={() => setIsAddingUser(!isAddingUser)}
                className="px-3.5 py-1.5 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-bold transition-all shadow-md shadow-purple-600/20 flex items-center gap-1.5 cursor-pointer"
              >
                <UserPlus className="w-3.5 h-3.5" />
                <span>Add User</span>
              </button>
            </div>
          </div>

          {/* Info panel: accounts under Entra ID sign-in can't be pre-created here */}
          {isAddingUser && (
            <motion.div
              initial={{ opacity: 0, height: 0 }}
              animate={{ opacity: 1, height: "auto" }}
              exit={{ opacity: 0, height: 0 }}
              className={`p-5 border-b space-y-2 shrink-0 font-sans ${
                isLightMode
                  ? "bg-purple-50/60 border-purple-200"
                  : "bg-purple-950/20 border-purple-500/30"
              }`}
            >
              <div className="flex items-center justify-between">
                <span className={`text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 ${
                  isLightMode ? "text-purple-800" : "text-purple-300"
                }`}>
                  <UserPlus className="w-4 h-4 text-purple-600 dark:text-purple-400" /> Adding Accounts
                </span>
                <button
                  type="button"
                  onClick={() => setIsAddingUser(false)}
                  className={isLightMode ? "text-slate-400 hover:text-slate-700 cursor-pointer" : "text-slate-400 hover:text-white cursor-pointer"}
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              <p className={`text-xs leading-relaxed ${isLightMode ? "text-slate-700" : "text-slate-400"}`}>
                Accounts are created automatically the first time someone signs in with
                their Halliburton Microsoft account — there's no manual creation step here.
                New sign-ins show up below under <span className="font-bold text-amber-400">Pending Access Requests</span>,
                where you can approve or decline them.
              </p>
            </motion.div>
          )}

          {/* User List */}
          <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-4 bg-[#121620]">
            {/* Separate Pending Section when viewing all or pending and pending users exist */}
            {statusFilter !== "approved" && pendingCount > 0 && !searchQuery && (
              <div className="p-4 rounded-2xl bg-amber-950/20 border-2 border-amber-500/40 shadow-lg shadow-amber-950/30 space-y-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <span className="relative flex h-2.5 w-2.5">
                      <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                      <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
                    </span>
                    <h4 className="text-xs font-bold uppercase tracking-wider text-amber-300 font-mono flex items-center gap-2">
                      <span>Pending Access Requests</span>
                      <span className="px-2 py-0.5 bg-amber-500/20 text-amber-200 rounded-full border border-amber-500/40 text-[10px] font-bold">
                        {pendingCount} awaiting approval
                      </span>
                    </h4>
                  </div>
                  <span className="text-[10px] text-amber-400/80 font-mono hidden sm:inline">
                    Access blocked until approved by Admin
                  </span>
                </div>

                <div className="space-y-2.5">
                  {users
                    .filter((u) => u.status === "pending")
                    .map((pu) => {
                      const pUid = pu.id || "";
                      return (
                        <div
                          key={`pending-sec-${pUid}`}
                          className="p-3.5 rounded-xl bg-amber-950/40 border border-amber-500/30 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                        >
                          <div className="flex items-center gap-3 min-w-0">
                            <div className="w-9 h-9 rounded-lg bg-amber-500/20 border border-amber-500/40 text-amber-300 flex items-center justify-center shrink-0">
                              <Clock className="w-4 h-4 animate-pulse text-amber-400" />
                            </div>
                            <div className="min-w-0">
                              <div className="flex items-center gap-2 flex-wrap">
                                <span className="font-bold text-xs text-white">{pu.name}</span>
                                <span className="text-[9px] font-extrabold uppercase px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-300 border border-amber-500/40 animate-pulse font-mono">
                                  Pending
                                </span>
                                {pu.requestedRole && (
                                  <span className="text-[9px] px-1.5 py-0.5 rounded bg-slate-800 text-slate-300 font-mono border border-slate-700">
                                    Req: {pu.requestedRole}
                                  </span>
                                )}
                              </div>
                              <div className="text-[11px] text-slate-400 font-mono truncate">
                                {pu.email} {pu.title ? `• ${pu.title}` : ""}
                              </div>
                            </div>
                          </div>

                          <div className="flex items-center gap-2 shrink-0">
                            <button
                              type="button"
                              onClick={() => handleUpdateUserQuick(pUid, { status: "approved" })}
                              className="px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs cursor-pointer flex items-center gap-1.5 shadow-md shadow-emerald-600/30 border border-emerald-400/40 transition-all active:scale-95"
                              id={`approve-pending-sec-${pUid}`}
                            >
                              <Check className="w-3.5 h-3.5" />
                              <span>Approve Access</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => handleUpdateUserQuick(pUid, { status: "rejected" })}
                              className="px-2.5 py-1.5 rounded-xl bg-rose-600/20 hover:bg-rose-600/40 text-rose-300 font-semibold text-xs cursor-pointer flex items-center gap-1 border border-rose-500/30 transition-all active:scale-95"
                              id={`decline-pending-sec-${pUid}`}
                            >
                              <X className="w-3.5 h-3.5" />
                              <span>Decline</span>
                            </button>
                          </div>
                        </div>
                      );
                    })}
                </div>
              </div>
            )}

            {filteredUsers.length === 0 ? (
              <div className="text-center py-12 text-xs text-slate-500 font-mono">
                {loading ? "Loading users from Firestore..." : "No user accounts match your search query."}
              </div>
            ) : (
              <div className="space-y-3">
                {filteredUsers.map((u) => {
                  const targetUid = u.id || "";
                  const isRootAdmin = u.email === "admin@drillops.com";
                  const isPending = u.status === "pending";

                  return (
                    <div
                      key={targetUid}
                      className={`p-4 rounded-2xl border transition-all flex flex-col md:flex-row md:items-center justify-between gap-4 ${
                        isPending
                          ? "bg-amber-950/25 border-amber-500/50 shadow-md shadow-amber-950/20"
                          : u.role === "admin"
                          ? "bg-purple-950/20 border-purple-500/30"
                          : u.role === "editor"
                          ? "bg-blue-950/20 border-blue-500/30"
                          : "bg-[#181E2E] border-slate-700/60"
                      }`}
                    >
                      {/* Left: User Identity Details */}
                      <div className="flex items-center gap-3 min-w-0">
                        <div
                          className={`w-10 h-10 rounded-xl flex items-center justify-center font-bold text-sm shrink-0 border ${
                            isPending
                              ? "bg-amber-600/20 text-amber-400 border-amber-500/40 animate-pulse"
                              : u.role === "admin"
                              ? "bg-purple-600/20 text-purple-400 border-purple-500/30"
                              : u.role === "editor"
                              ? "bg-blue-600/20 text-blue-400 border-blue-500/30"
                              : "bg-emerald-600/20 text-emerald-400 border-emerald-500/30"
                          }`}
                        >
                          {isPending ? (
                            <Clock className="w-5 h-5 text-amber-400" />
                          ) : u.role === "admin" ? (
                            <Shield className="w-5 h-5" />
                          ) : u.role === "editor" ? (
                            <ShieldCheck className="w-5 h-5" />
                          ) : (
                            <Eye className="w-5 h-5" />
                          )}
                        </div>

                        <div className="min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-bold text-sm text-slate-100">{u.name}</span>

                            <span
                              className={`text-[9px] font-extrabold uppercase px-2 py-0.5 rounded ${
                                u.role === "admin"
                                  ? "bg-purple-500/20 text-purple-300 border border-purple-500/30"
                                  : u.role === "editor"
                                  ? "bg-blue-500/20 text-blue-300 border border-blue-500/30"
                                  : "bg-slate-500/20 text-slate-300 border border-slate-500/30"
                              }`}
                            >
                              {u.role === "admin" ? "Admin" : u.role === "editor" ? "Editor" : "Viewer"}
                            </span>

                            {/* Clear Visual Indicator for User Status */}
                            <span
                              className={`text-[9px] font-extrabold uppercase px-2.5 py-0.5 rounded-full flex items-center gap-1 ${
                                u.status === "approved"
                                  ? "bg-emerald-500/20 text-emerald-300 border border-emerald-500/40"
                                  : u.status === "rejected"
                                  ? "bg-rose-500/20 text-rose-300 border border-rose-500/40"
                                  : "bg-amber-500/25 text-amber-300 border border-amber-500/50 animate-pulse font-mono"
                              }`}
                            >
                              {isPending && <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-ping" />}
                              <span>{u.status === "approved" ? "Approved" : u.status === "rejected" ? "Rejected" : "Pending"}</span>
                            </span>
                          </div>

                          <div className="flex items-center gap-3 text-xs mt-1 font-mono flex-wrap text-slate-400">
                            <span>Email: <strong className="text-slate-200">{u.email}</strong></span>
                            {u.title && <span>• <span className="text-slate-300">{u.title}</span></span>}
                          </div>
                        </div>
                      </div>

                      {/* Right: Quick Actions & Role Switcher */}
                      <div className="flex items-center gap-2 flex-wrap shrink-0">
                        {/* Role Selector Buttons */}
                        <div className="flex items-center gap-1 border border-slate-700/60 rounded-xl p-1 text-[11px] bg-[#0E121B]">
                          <button
                            type="button"
                            onClick={() =>
                              handleUpdateUserQuick(targetUid, { role: "admin", permissions: "both" })
                            }
                            className={`px-2.5 py-1 rounded-lg font-bold transition-all cursor-pointer ${
                              u.role === "admin"
                                ? "bg-purple-600 text-white shadow"
                                : "text-slate-400 hover:text-white"
                            }`}
                          >
                            Admin
                          </button>
                          <button
                            type="button"
                            onClick={() =>
                              handleUpdateUserQuick(targetUid, { role: "editor", permissions: "both" })
                            }
                            className={`px-2.5 py-1 rounded-lg font-bold transition-all cursor-pointer ${
                              u.role === "editor"
                                ? "bg-blue-600 text-white shadow"
                                : "text-slate-400 hover:text-white"
                            }`}
                          >
                            Editor
                          </button>
                          <button
                            type="button"
                            onClick={() =>
                              handleUpdateUserQuick(targetUid, { role: "viewer", permissions: "read" })
                            }
                            className={`px-2.5 py-1 rounded-lg font-bold transition-all cursor-pointer ${
                              u.role === "viewer"
                                ? "bg-amber-600 text-white shadow"
                                : "text-slate-400 hover:text-white"
                            }`}
                          >
                            Viewer
                          </button>
                        </div>

                        {/* Approve Button for Pending Users */}
                        {isPending && (
                          <button
                            type="button"
                            onClick={() => handleUpdateUserQuick(targetUid, { status: "approved" })}
                            className="px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs cursor-pointer flex items-center gap-1.5 shadow-md shadow-emerald-600/30 border border-emerald-400/40 transition-all active:scale-95"
                            title="Approve access for this user"
                            id={`approve-user-btn-${targetUid}`}
                          >
                            <Check className="w-3.5 h-3.5" />
                            <span>Approve</span>
                          </button>
                        )}

                        {/* Revoke Button next to existing role controls to set approved user back to pending */}
                        {!isPending && !isRootAdmin && targetUid !== currentUser.id && (
                          <button
                            type="button"
                            onClick={() => handleUpdateUserQuick(targetUid, { status: "pending" })}
                            className="px-2.5 py-1.5 rounded-xl text-xs font-semibold transition-all cursor-pointer border border-amber-500/40 bg-amber-950/30 hover:bg-amber-900/60 active:scale-95 text-amber-300 hover:text-amber-200 flex items-center gap-1 shadow-sm"
                            title="Revoke access (set status back to pending)"
                            id={`revoke-user-btn-${targetUid}`}
                          >
                            <UserX className="w-3.5 h-3.5 text-amber-400" />
                            <span>Revoke</span>
                          </button>
                        )}

                        {/* Edit Account Modal trigger */}
                        <button
                          type="button"
                          onClick={() => openEditModal(u)}
                          className="px-2.5 py-1.5 rounded-xl text-xs font-semibold transition-all cursor-pointer border bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700 flex items-center gap-1"
                        >
                          <Edit3 className="w-3.5 h-3.5 text-blue-400" />
                          <span>Edit</span>
                        </button>

                        {/* Delete User button */}
                        {!isRootAdmin && targetUid !== currentUser.id && (
                          <button
                            type="button"
                            onClick={() => handleDeleteUser(targetUid, u.name, u.email)}
                            className="p-2 rounded-xl border border-slate-800 transition-colors cursor-pointer bg-slate-900 text-slate-500 hover:text-rose-400 hover:border-rose-500/30"
                            title="Delete User Account"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Edit User Modal */}
          {editingUser && (
            <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center z-[120] p-4">
              <motion.form
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                onSubmit={handleSaveEditUser}
                className={`border rounded-2xl p-6 max-w-md w-full space-y-4 shadow-2xl ${
                  isLightMode
                    ? "bg-slate-200 border-slate-300 text-slate-900"
                    : "bg-[#141A29] border-slate-700 text-slate-100"
                }`}
              >
                <div className={`flex items-center justify-between border-b pb-3 ${
                  isLightMode ? "border-slate-300" : "border-slate-800"
                }`}>
                  <div className="flex items-center gap-2">
                    <Edit3 className={`w-5 h-5 ${isLightMode ? "text-blue-600" : "text-blue-400"}`} />
                    <h4 className="font-bold text-sm">Edit User: {editingUser.name}</h4>
                  </div>
                  <button
                    type="button"
                    onClick={() => setEditingUser(null)}
                    className={isLightMode ? "text-slate-400 hover:text-slate-800 cursor-pointer" : "text-slate-400 hover:text-white cursor-pointer"}
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>

                <div className="space-y-3">
                  <div>
                    <label className={`block text-xs font-semibold mb-1 ${isLightMode ? "text-slate-700" : "text-slate-400"}`}>Account Name</label>
                    <input
                      type="text"
                      required
                      value={editName}
                      onChange={(e) => setEditName(e.target.value)}
                      className={`w-full h-9 px-3 text-xs border rounded-xl focus:outline-none focus:border-blue-500 ${
                        isLightMode
                          ? "bg-slate-50 border-slate-300 text-slate-900"
                          : "bg-[#0B0E17] border-slate-800 text-slate-200"
                      }`}
                    />
                  </div>

                  <div>
                    <label className={`block text-xs font-semibold mb-1 ${isLightMode ? "text-slate-700" : "text-slate-400"}`}>Email</label>
                    <input
                      type="email"
                      required
                      value={editEmail}
                      onChange={(e) => setEditEmail(e.target.value)}
                      className={`w-full h-9 px-3 text-xs border rounded-xl font-mono focus:outline-none focus:border-blue-500 ${
                        isLightMode
                          ? "bg-slate-50 border-slate-300 text-slate-900"
                          : "bg-[#0B0E17] border-slate-800 text-slate-200"
                      }`}
                    />
                  </div>

                  <div>
                    <label className={`block text-xs font-semibold mb-1 ${isLightMode ? "text-slate-700" : "text-slate-400"}`}>Approval Status</label>
                    <select
                      value={editStatus}
                      onChange={(e) => setEditStatus(e.target.value as UserStatus)}
                      className={`w-full h-9 px-3 text-xs border rounded-xl font-semibold focus:outline-none focus:border-blue-500 ${
                        isLightMode
                          ? "bg-slate-50 border-slate-300 text-slate-900"
                          : "bg-[#0B0E17] border-slate-800 text-slate-200"
                      }`}
                    >
                      <option value="approved">Approved</option>
                      <option value="pending">Pending Approval</option>
                      <option value="rejected">Rejected</option>
                    </select>
                  </div>

                  <div>
                    <label className={`block text-xs font-semibold mb-1 ${isLightMode ? "text-slate-700" : "text-slate-400"}`}>Role & Access</label>
                    <select
                      value={editRole}
                      onChange={(e) => {
                        const r = e.target.value as UserRole;
                        setEditRole(r);
                        if (r === "admin" || r === "editor") setEditPermissions("both");
                        else setEditPermissions("read");
                      }}
                      className={`w-full h-9 px-3 text-xs border rounded-xl font-semibold focus:outline-none focus:border-blue-500 ${
                        isLightMode
                          ? "bg-slate-50 border-slate-300 text-slate-900"
                          : "bg-[#0B0E17] border-slate-800 text-slate-200"
                      }`}
                    >
                      <option value="viewer">User (Read-Only)</option>
                      <option value="editor">Editor (Write Rights)</option>
                      <option value="admin">Admin (Full Control)</option>
                    </select>
                  </div>
                </div>

                <div className={`flex items-center justify-end gap-2 pt-2 border-t ${
                  isLightMode ? "border-slate-200" : "border-slate-800"
                }`}>
                  <button
                    type="button"
                    onClick={() => setEditingUser(null)}
                    className={`px-3 py-1.5 rounded-xl text-xs cursor-pointer ${
                      isLightMode
                        ? "bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold"
                        : "bg-slate-800 hover:bg-slate-700 text-slate-300"
                    }`}
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-1.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs shadow-md shadow-blue-600/20 cursor-pointer"
                  >
                    Save Changes
                  </button>
                </div>
              </motion.form>
            </div>
          )}
        </motion.div>
      </div>
    </AnimatePresence>
  );
}
