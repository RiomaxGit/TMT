import { useState, useRef, useCallback, useMemo, useEffect, useReducer } from "react";
import * as XLSX from "xlsx";

// ─── Google Sheets Config ─────────────────────────────────────────────────────
const GAPI_CLIENT_ID = "978072122166-q1mekaie0lhh0lsg7vr0rtpo12mfp9e9.apps.googleusercontent.com";
const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";
const SHEET_ID_KEY = "testflow_sheet_id";
const SHEET_TABS = ["Projects", "Modules", "TestCases", "Suites", "Runs"];

// ─── Theme ────────────────────────────────────────────────────────────────────
function makeTheme(dark) {
  return dark ? {
    bg: "#0F1117", surface: "#1A1D27", surface2: "#22263A", surface3: "#2A2F45",
    border: "#2E3347", borderStrong: "#3D4460",
    text: "#F1F5F9", textSec: "#94A3B8", textMuted: "#5A6478",
    blue: "#4B8EF0", blueSoft: "#1E2D4A", blueText: "#93BBFF",
    green: "#34D399", greenSoft: "#0D2B1F", greenText: "#6EE7B7",
    red: "#F87171", redSoft: "#2D1515", redText: "#FCA5A5",
    amber: "#FBBF24", amberSoft: "#2D2009", amberText: "#FDE68A",
    purple: "#A78BFA", purpleSoft: "#1F1535", purpleText: "#C4B5FD",
    teal: "#2DD4BF", tealSoft: "#0D2926",
    gray100: "#1E2132", gray200: "#252A3D",
    inputBg: "#1E2132", shadow: "rgba(0,0,0,0.5)",
    navActive: "#1E2D4A", isDark: true,
  } : {
    bg: "#F4F6FB", surface: "#FFFFFF", surface2: "#F9FAFB", surface3: "#F1F5F9",
    border: "#E5E7EB", borderStrong: "#D1D5DB",
    text: "#111827", textSec: "#4B5563", textMuted: "#9CA3AF",
    blue: "#1A6BCC", blueSoft: "#EBF2FF", blueText: "#1559AB",
    green: "#059669", greenSoft: "#ECFDF5", greenText: "#047857",
    red: "#DC2626", redSoft: "#FEF2F2", redText: "#B91C1C",
    amber: "#D97706", amberSoft: "#FFFBEB", amberText: "#B45309",
    purple: "#7C3AED", purpleSoft: "#F5F3FF", purpleText: "#6D28D9",
    teal: "#0D9488", tealSoft: "#F0FDFA",
    gray100: "#F3F4F6", gray200: "#E5E7EB",
    inputBg: "#FFFFFF", shadow: "rgba(0,0,0,0.12)",
    navActive: "#EBF2FF", isDark: false,
  };
}

const STATUS_META = {
  pass:         { label: "Pass",        colorKey: "green",    icon: "✓" },
  fail:         { label: "Fail",        colorKey: "red",      icon: "✗" },
  pending:      { label: "Pending",     colorKey: "textMuted",icon: "○" },
  blocked:      { label: "Blocked",     colorKey: "amber",    icon: "⊘" },
  skipped:      { label: "Skipped",     colorKey: "purple",   icon: "↷" },
  "in-progress":{ label: "In Progress", colorKey: "blue",     icon: "◐" },
};

// ─── Utilities ────────────────────────────────────────────────────────────────
function uid() { return Math.random().toString(36).slice(2, 10); }
function genTCId(pk, mk, n) { return `${pk}-${mk}-TC${String(n).padStart(3, "0")}`; }

// ─── Primitive UI Components ──────────────────────────────────────────────────
function StatusBadge({ status, T, small }) {
  const m = STATUS_META[status] || STATUS_META.pending;
  const color = T[m.colorKey] || T.textMuted;
  const bg = T[m.colorKey + "Soft"] || T.gray100;
  return (
    <span style={{ display:"inline-flex", alignItems:"center", gap:3, padding:small?"2px 7px":"3px 10px", borderRadius:20, fontSize:small?10:12, fontWeight:600, color, background:bg, whiteSpace:"nowrap" }}>
      {m.icon} {m.label}
    </span>
  );
}

function PriBadge({ p, T }) {
  const meta = { high:{ c:T.red, bg:T.redSoft }, medium:{ c:T.amber, bg:T.amberSoft }, low:{ c:T.green, bg:T.greenSoft } };
  const m = meta[p] || meta.medium;
  return <span style={{ fontSize:10, fontWeight:700, color:m.c, background:m.bg, padding:"1px 7px", borderRadius:20, textTransform:"uppercase", letterSpacing:0.5 }}>{p}</span>;
}

function RunStatusBadge({ status, T }) {
  const meta = { active:{ c:T.blue, bg:T.blueSoft }, completed:{ c:T.green, bg:T.greenSoft }, draft:{ c:T.textMuted, bg:T.gray100 } };
  const m = meta[status] || meta.draft;
  return <span style={{ fontSize:11, fontWeight:600, color:m.c, background:m.bg, padding:"2px 8px", borderRadius:20 }}>{status === "active" ? "Active" : status === "completed" ? "Completed" : "Draft"}</span>;
}

function Btn({ onClick, variant = "ghost", children, style, disabled, title, T }) {
  const base = { border:"1px solid", borderRadius:8, cursor:disabled?"not-allowed":"pointer", fontFamily:"inherit", fontSize:13, fontWeight:500, display:"inline-flex", alignItems:"center", gap:6, padding:"6px 14px", transition:"all 0.13s", opacity:disabled?0.45:1, outline:"none" };
  const vv = {
    primary: { background:T.blue,    color:"#fff", borderColor:T.blue },
    danger:  { background:T.red,     color:"#fff", borderColor:T.red },
    success: { background:T.green,   color:"#fff", borderColor:T.green },
    outline: { background:T.surface, color:T.text, borderColor:T.border },
    ghost:   { background:"transparent", color:T.textSec, borderColor:"transparent" },
    warning: { background:T.amber,   color:"#fff", borderColor:T.amber },
  };
  return <button onClick={disabled ? undefined : onClick} title={title} style={{ ...base, ...vv[variant], ...style }}>{children}</button>;
}

function Modal({ title, onClose, children, width = 560, T }) {
  return (
    <div style={{ position:"fixed", inset:0, zIndex:1000, background:"rgba(0,0,0,0.55)", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}
      onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="fade-in" style={{ background:T.surface, borderRadius:16, width:"100%", maxWidth:width, maxHeight:"92vh", overflow:"auto", boxShadow:`0 24px 80px ${T.shadow}`, border:`1px solid ${T.border}` }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"18px 24px", borderBottom:`1px solid ${T.border}`, position:"sticky", top:0, background:T.surface, zIndex:1 }}>
          <h2 style={{ margin:0, fontSize:16, fontWeight:700, color:T.text }}>{title}</h2>
          <Btn onClick={onClose} variant="ghost" T={T} style={{ padding:"4px 8px", color:T.textMuted }}>✕</Btn>
        </div>
        <div style={{ padding:"20px 24px" }}>{children}</div>
      </div>
    </div>
  );
}

function ConfirmModal({ title, message, confirmLabel = "Delete", onConfirm, onClose, T }) {
  return (
    <Modal title={title} onClose={onClose} width={420} T={T}>
      <p style={{ color:T.textSec, fontSize:14, lineHeight:1.6, marginBottom:20 }}>{message}</p>
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Cancel</Btn>
        <Btn onClick={() => { onConfirm(); onClose(); }} variant="danger" T={T}>{confirmLabel}</Btn>
      </div>
    </Modal>
  );
}

function Field({ label, children, required, T }) {
  return (
    <div style={{ marginBottom:16 }}>
      <label style={{ display:"block", fontSize:13, fontWeight:500, color:T.textSec, marginBottom:6 }}>
        {label}{required && <span style={{ color:T.red }}> *</span>}
      </label>
      {children}
    </div>
  );
}

function iStyle(T) {
  return { width:"100%", boxSizing:"border-box", border:`1px solid ${T.border}`, borderRadius:8, padding:"8px 12px", fontSize:13, fontFamily:"inherit", color:T.text, background:T.inputBg, outline:"none" };
}

function TfInput({ T, ...props }) { return <input style={iStyle(T)} {...props} />; }
function TfTextarea({ T, minH = 80, ...props }) { return <textarea style={{ ...iStyle(T), resize:"vertical", minHeight:minH }} {...props} />; }
function TfSelect({ T, children, ...props }) { return <select style={{ ...iStyle(T), cursor:"pointer" }} {...props}>{children}</select>; }

function SearchBar({ value, onChange, placeholder, T }) {
  return (
    <div style={{ position:"relative" }}>
      <span style={{ position:"absolute", left:10, top:"50%", transform:"translateY(-50%)", color:T.textMuted, fontSize:13, pointerEvents:"none" }}>🔍</span>
      <input value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
        style={{ ...iStyle(T), paddingLeft:32, fontSize:12 }} />
    </div>
  );
}

// ─── Toast Notification ───────────────────────────────────────────────────────
function Toast({ message, type, T }) {
  const colors = { success:T.green, error:T.red, info:T.blue, warning:T.amber };
  const bgs = { success:T.greenSoft, error:T.redSoft, info:T.blueSoft, warning:T.amberSoft };
  return (
    <div className="fade-in" style={{ position:"fixed", bottom:24, right:24, zIndex:9999, background:bgs[type]||T.surface, border:`1px solid ${colors[type]||T.border}`, borderRadius:12, padding:"12px 18px", display:"flex", alignItems:"center", gap:10, boxShadow:`0 8px 30px ${T.shadow}`, maxWidth:340 }}>
      <span style={{ fontSize:16 }}>{type==="success"?"✅":type==="error"?"❌":type==="warning"?"⚠️":"ℹ️"}</span>
      <span style={{ fontSize:13, color:T.text, fontWeight:500 }}>{message}</span>
    </div>
  );
}

// ─── Google Sheets Integration ────────────────────────────────────────────────

/**
 * useGoogleSheets — full lifecycle hook
 *
 * FLOW:
 *  1. On mount, injects the GIS script if not already present.
 *  2. signIn()  → opens OAuth popup → stores access token in memory.
 *  3. saveToSheets(state) → creates the spreadsheet on first save (stores ID in
 *     localStorage under SHEET_ID_KEY), then batch-writes all 5 tabs.
 *  4. loadFromSheets() → reads all 5 tabs from the stored sheet and returns
 *     a parsed state object that callers dispatch as IMPORT.
 *  5. signOut() → revokes the token + clears stored sheet id.
 */
function useGoogleSheets(onToast) {
  const [gsUser, setGsUser]     = useState(null);   // { name, email, picture }
  const [gsStatus, setGsStatus] = useState("idle"); // idle | signing-in | ready | saving | loading | error
  const tokenRef = useRef(null);
  const clientRef = useRef(null);

  // ---------- helpers ----------
  const apiFetch = useCallback(async (url, options = {}) => {
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${tokenRef.current}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: { message: res.statusText } }));
      throw new Error(err?.error?.message || `HTTP ${res.status}`);
    }
    return res.json();
  }, []);

  const getUserInfo = useCallback(async () => {
    const data = await apiFetch("https://www.googleapis.com/oauth2/v3/userinfo");
    return { name: data.name, email: data.email, picture: data.picture };
  }, [apiFetch]);

  // ---------- init GIS ----------
  useEffect(() => {
    const initClient = () => {
      clientRef.current = window.google.accounts.oauth2.initTokenClient({
        client_id: GAPI_CLIENT_ID,
        scope: SCOPES,
        callback: async (resp) => {
          if (resp.error) {
            setGsStatus("error");
            onToast("Google sign-in failed: " + resp.error, "error");
            return;
          }
          tokenRef.current = resp.access_token;
          try {
            const user = await getUserInfo();
            setGsUser(user);
            setGsStatus("ready");
            onToast(`Signed in as ${user.email}`, "success");
          } catch (e) {
            setGsStatus("error");
            onToast("Could not fetch user info", "error");
          }
        },
      });
    };

    if (window.google?.accounts?.oauth2) {
      initClient();
    } else {
      const script = document.createElement("script");
      script.src = "https://accounts.google.com/gsi/client";
      script.onload = initClient;
      document.head.appendChild(script);
    }
  }, [getUserInfo, onToast]);

  // ---------- sign in ----------
  const signIn = useCallback(() => {
    if (!clientRef.current) { onToast("Google API not loaded yet, try again in a moment", "warning"); return; }
    setGsStatus("signing-in");
    clientRef.current.requestAccessToken({ prompt: "consent" });
  }, [onToast]);

  // ---------- sign out ----------
  const signOut = useCallback(() => {
    if (tokenRef.current) window.google?.accounts?.oauth2?.revoke(tokenRef.current);
    tokenRef.current = null;
    setGsUser(null);
    setGsStatus("idle");
    onToast("Signed out from Google", "info");
  }, [onToast]);

  // ---------- serialisers ----------
  const stateToSheets = useCallback((state) => {
    const projects = [
      ["ID","Name","Key","Description","CreatedAt"],
      ...state.projects.map(p => [p.id, p.name, p.key, p.description, p.createdAt]),
    ];
    const modules = [
      ["ID","ProjectID","Name","Key","Description"],
      ...state.modules.map(m => [m.id, m.projectId, m.name, m.key, m.description]),
    ];
    const testcases = [
      ["ID","ProjectID","ModuleID","TCID","Num","Title","Description","Steps","ExpectedResult","Priority","Tags","Status"],
      ...state.testcases.map(tc => [tc.id, tc.projectId, tc.moduleId, tc.tcId, tc.num, tc.title, tc.description, tc.steps, tc.expectedResult, tc.priority, tc.tags, tc.status]),
    ];
    const suites = [
      ["ID","Name","Description","CreatedAt","Inclusions"],
      ...state.suites.map(s => [s.id, s.name, s.description, s.createdAt, JSON.stringify(s.inclusions)]),
    ];
    const runRows = [["RunID","RunName","RunStatus","CreatedAt","CompletedAt","TCID","CaseStatus","Comment","ExecutedAt"]];
    state.runs.forEach(r => r.cases.forEach(c => runRows.push([r.id, r.name, r.status, r.createdAt, r.completedAt || "", c.tcId, c.status, c.comment, c.executedAt || ""])));
    if (runRows.length === 1) runRows.push(["","","","","","","","",""]);
    return { Projects: projects, Modules: modules, TestCases: testcases, Suites: suites, Runs: runRows };
  }, []);

  const sheetsToState = useCallback((rawTabs) => {
    const parse = (rows) => {
      if (!rows || rows.length < 2) return [];
      const [headers, ...data] = rows;
      return data.map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i] ?? ""; });
        return obj;
      });
    };
    const projects  = parse(rawTabs.Projects).filter(r => r.ID).map(r => ({ id:r.ID, name:r.Name, key:r.Key, description:r.Description, createdAt:r.CreatedAt }));
    const modules   = parse(rawTabs.Modules).filter(r => r.ID).map(r => ({ id:r.ID, projectId:r.ProjectID, name:r.Name, key:r.Key, description:r.Description }));
    const testcases = parse(rawTabs.TestCases).filter(r => r.ID).map(r => ({ id:r.ID, projectId:r.ProjectID, moduleId:r.ModuleID, tcId:r.TCID, num:Number(r.Num)||1, title:r.Title, description:r.Description, steps:r.Steps, expectedResult:r.ExpectedResult, priority:r.Priority||"medium", tags:r.Tags, status:r.Status||"active" }));
    const suites    = parse(rawTabs.Suites).filter(r => r.ID).map(r => ({ id:r.ID, name:r.Name, description:r.Description, createdAt:r.CreatedAt, inclusions: (() => { try { return JSON.parse(r.Inclusions||"[]"); } catch { return []; } })() }));
    const runMap = {};
    parse(rawTabs.Runs).filter(r => r.RunID).forEach(r => {
      if (!runMap[r.RunID]) runMap[r.RunID] = { id:r.RunID, name:r.RunName, description:"", suiteIds:[], cases:[], status:r.RunStatus||"active", createdAt:r.CreatedAt, completedAt:r.CompletedAt||null };
      if (r.TCID) runMap[r.RunID].cases.push({ tcId:r.TCID, status:r.CaseStatus||"pending", comment:r.Comment||"", executedAt:r.ExecutedAt||null });
    });
    return { projects, modules, testcases, suites, runs: Object.values(runMap) };
  }, []);

  // ---------- ensure sheet tabs exist ----------
  const ensureTabs = useCallback(async (spreadsheetId, existingSheets) => {
    const existingTitles = existingSheets.map(s => s.properties.title);
    const missing = SHEET_TABS.filter(t => !existingTitles.includes(t));
    if (!missing.length) return;
    await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
      method: "POST",
      body: JSON.stringify({ requests: missing.map(title => ({ addSheet: { properties: { title } } })) }),
    });
  }, [apiFetch]);

  // ---------- save ----------
  const saveToSheets = useCallback(async (state) => {
    if (!tokenRef.current) { onToast("Please sign in to Google first", "warning"); return; }
    setGsStatus("saving");
    try {
      let sheetId = localStorage.getItem(SHEET_ID_KEY);

      // Create spreadsheet if first time
      if (!sheetId) {
        const created = await apiFetch("https://sheets.googleapis.com/v4/spreadsheets", {
          method: "POST",
          body: JSON.stringify({ properties: { title: "TestFlow Data" }, sheets: SHEET_TABS.map(title => ({ properties: { title } })) }),
        });
        sheetId = created.spreadsheetId;
        localStorage.setItem(SHEET_ID_KEY, sheetId);
        onToast("Created new Google Sheet for TestFlow ✓", "success");
      } else {
        // Make sure all tabs exist (in case someone deleted one)
        const meta = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties`);
        await ensureTabs(sheetId, meta.sheets || []);
      }

      const sheets = stateToSheets(state);

      // Clear then write each tab
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchClear`, {
        method: "POST",
        body: JSON.stringify({ ranges: SHEET_TABS.map(t => `${t}!A:Z`) }),
      });
      // await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate`, {
      //   method: "PUT",
      //   body: JSON.stringify({
      //     valueInputOption: "RAW",
      //     data: SHEET_TABS.map(tab => ({ range: `${tab}!A1`, values: sheets[tab] })),
      //   }),
      // });
      await apiFetch(
  `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchUpdate?valueInputOption=RAW`,
  {
    method: "POST",
    body: JSON.stringify({
      data: SHEET_TABS.map(tab => ({
        range: `${tab}!A1`,
        values: sheets[tab],
      })),
    }),
  }
);

      const sheetUrl = `https://docs.google.com/spreadsheets/d/${sheetId}`;
      onToast("Saved to Google Sheets ✓", "success");
      setGsStatus("ready");
      return sheetUrl;
    } catch (e) {
      setGsStatus("ready");
      onToast("Save failed: " + e.message, "error");
    }
  }, [apiFetch, ensureTabs, onToast, stateToSheets]);

  // ---------- load ----------
  const loadFromSheets = useCallback(async () => {
    if (!tokenRef.current) { onToast("Please sign in to Google first", "warning"); return null; }
    const sheetId = localStorage.getItem(SHEET_ID_KEY);
    if (!sheetId) { onToast("No saved sheet found. Save first to create one.", "warning"); return null; }
    setGsStatus("loading");
    try {
      const ranges = SHEET_TABS.map(t => `${t}!A:Z`).join("&ranges=");
      const data = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values:batchGet?ranges=${ranges}`);
      const rawTabs = {};
      (data.valueRanges || []).forEach((vr, i) => { rawTabs[SHEET_TABS[i]] = vr.values || []; });
      const parsed = sheetsToState(rawTabs);
      setGsStatus("ready");
      onToast("Loaded from Google Sheets ✓", "success");
      return parsed;
    } catch (e) {
      setGsStatus("ready");
      onToast("Load failed: " + e.message, "error");
      return null;
    }
  }, [apiFetch, onToast, sheetsToState]);

  // ---------- open sheet in browser ----------
  const openSheet = useCallback(() => {
    const sheetId = localStorage.getItem(SHEET_ID_KEY);
    if (sheetId) window.open(`https://docs.google.com/spreadsheets/d/${sheetId}`, "_blank");
    else onToast("No sheet yet — save first", "warning");
  }, [onToast]);

  return { gsUser, gsStatus, signIn, signOut, saveToSheets, loadFromSheets, openSheet };
}

// ─── Initial Data ─────────────────────────────────────────────────────────────
function makeInitial() {
  const p1 = uid(), p2 = uid(), m1 = uid(), m2 = uid(), m3 = uid();
  const t1 = uid(), t2 = uid(), t3 = uid(), t4 = uid();
  return {
    projects: [
      { id:p1, name:"E-Commerce Platform", key:"ECP", description:"Online shopping platform", createdAt:new Date().toISOString() },
      { id:p2, name:"Admin Dashboard",     key:"ADM", description:"Internal admin tool",      createdAt:new Date().toISOString() },
    ],
    modules: [
      { id:m1, projectId:p1, name:"Authentication",  key:"AUTH", description:"Login, register, session" },
      { id:m2, projectId:p1, name:"Cart & Checkout",  key:"CART", description:"Shopping cart and payment" },
      { id:m3, projectId:p2, name:"User Management",  key:"USR",  description:"CRUD for users" },
    ],
    testcases: [
      { id:t1, moduleId:m1, projectId:p1, tcId:"ECP-AUTH-TC001", title:"Valid login with correct credentials",   description:"Verify user can login with correct email and password", steps:"1. Navigate to /login\n2. Enter valid email\n3. Enter valid password\n4. Click Sign In",      expectedResult:"User is redirected to dashboard. Session token is set.", priority:"high",   status:"active", num:1, tags:"smoke,regression" },
      { id:t2, moduleId:m1, projectId:p1, tcId:"ECP-AUTH-TC002", title:"Login with invalid password shows error", description:"Verify appropriate error for wrong password",            steps:"1. Navigate to /login\n2. Enter valid email\n3. Enter wrong password\n4. Click Sign In", expectedResult:"Error message: 'Invalid credentials'. User stays on login page.", priority:"medium", status:"active", num:2, tags:"regression" },
      { id:t3, moduleId:m2, projectId:p1, tcId:"ECP-CART-TC001", title:"Add item to cart",                       description:"Verify product can be added to shopping cart",           steps:"1. Browse to product page\n2. Click 'Add to Cart'\n3. View cart",                       expectedResult:"Item appears in cart with correct price and quantity.",  priority:"high",   status:"active", num:1, tags:"smoke" },
      { id:t4, moduleId:m3, projectId:p2, tcId:"ADM-USR-TC001",  title:"Create new user account",                description:"Admin can create a new user from the management panel", steps:"1. Login as admin\n2. Go to Users\n3. Click Create User\n4. Fill form\n5. Submit",       expectedResult:"New user appears in list. Welcome email sent.",          priority:"high",   status:"active", num:1, tags:"smoke" },
    ],
    suites: [],
    runs:   [],
  };
}

function resolveSuiteTCs(suite, state) {
  const set = new Set();
  suite.inclusions.forEach(inc => {
    if (inc.type === "project") state.testcases.filter(tc => tc.projectId === inc.id).forEach(tc => set.add(tc.id));
    else if (inc.type === "module") state.testcases.filter(tc => tc.moduleId === inc.id).forEach(tc => set.add(tc.id));
    else if (inc.type === "tc") set.add(inc.id);
  });
  return state.testcases.filter(tc => set.has(tc.id));
}

// ─── Reducer ──────────────────────────────────────────────────────────────────
function reducer(state, action) {
  switch (action.type) {
    case "ADD_PROJECT":    return { ...state, projects: [...state.projects, action.project] };
    case "EDIT_PROJECT":   return { ...state, projects: state.projects.map(p => p.id === action.project.id ? action.project : p) };
    case "DELETE_PROJECT": {
      const modIds = state.modules.filter(m => m.projectId === action.id).map(m => m.id);
      return {
        ...state,
        projects:  state.projects.filter(p => p.id !== action.id),
        modules:   state.modules.filter(m => m.projectId !== action.id),
        testcases: state.testcases.filter(tc => tc.projectId !== action.id),
        suites:    state.suites.map(s => ({ ...s, inclusions: s.inclusions.filter(i => i.id !== action.id && !modIds.includes(i.id)) })),
        runs:      state.runs.map(r => ({ ...r, cases: r.cases.filter(c => { const tc = state.testcases.find(t => t.id === c.tcId); return tc && tc.projectId !== action.id; }) })),
      };
    }
    case "ADD_MODULE":    return { ...state, modules: [...state.modules, action.module] };
    case "EDIT_MODULE":   return { ...state, modules: state.modules.map(m => m.id === action.module.id ? action.module : m) };
    case "DELETE_MODULE": return {
      ...state,
      modules:   state.modules.filter(m => m.id !== action.id),
      testcases: state.testcases.filter(tc => tc.moduleId !== action.id),
      suites:    state.suites.map(s => ({ ...s, inclusions: s.inclusions.filter(i => i.id !== action.id) })),
      runs:      state.runs.map(r => ({ ...r, cases: r.cases.filter(c => { const tc = state.testcases.find(t => t.id === c.tcId); return tc && tc.moduleId !== action.id; }) })),
    };
    case "ADD_TC":    return { ...state, testcases: [...state.testcases, action.tc] };
    case "EDIT_TC":   return { ...state, testcases: state.testcases.map(tc => tc.id === action.tc.id ? action.tc : tc) };
    case "DELETE_TC": return { ...state, testcases: state.testcases.filter(tc => tc.id !== action.id) };
    case "ADD_SUITE":    return { ...state, suites: [...state.suites, action.suite] };
    case "EDIT_SUITE":   return { ...state, suites: state.suites.map(s => s.id === action.suite.id ? action.suite : s) };
    case "DELETE_SUITE": return { ...state, suites: state.suites.filter(s => s.id !== action.id) };
    case "ADD_RUN":    return { ...state, runs: [...state.runs, action.run] };
    case "DELETE_RUN": return { ...state, runs: state.runs.filter(r => r.id !== action.id) };
    case "UPDATE_RUN_CASE": return { ...state, runs: state.runs.map(r => r.id === action.runId ? { ...r, cases: r.cases.map(c => c.tcId === action.tcId ? { ...c, ...action.update } : c) } : r) };
    case "COMPLETE_RUN":    return { ...state, runs: state.runs.map(r => r.id === action.runId ? { ...r, status:"completed", completedAt:new Date().toISOString() } : r) };
    case "IMPORT": return { ...state, ...action.data };
    default: return state;
  }
}

// ─── Project / Module / TC Modals ─────────────────────────────────────────────
function AddEditProjectModal({ existing, onClose, onSave, T }) {
  const [f, sf] = useState(existing || { name:"", key:"", description:"" });
  const ch = k => e => sf(p => ({ ...p, [k]: e.target.value }));
  return (
    <Modal title={existing ? "Edit Project" : "New Project"} onClose={onClose} width={480} T={T}>
      <Field label="Project Name" required T={T}><TfInput T={T} value={f.name} onChange={ch("name")} placeholder="E.g. Mobile App" /></Field>
      <Field label="Key (short prefix)" required T={T}><TfInput T={T} value={f.key} onChange={ch("key")} placeholder="E.g. APP" maxLength={6} /></Field>
      <Field label="Description" T={T}><TfTextarea T={T} value={f.description} onChange={ch("description")} /></Field>
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8, marginTop:8 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Cancel</Btn>
        <Btn variant="primary" T={T} onClick={() => { if (!f.name || !f.key) return; onSave({ ...(existing || {}), id:existing?.id || uid(), ...f, key:f.key.toUpperCase(), createdAt:existing?.createdAt || new Date().toISOString() }); }}>
          {existing ? "Save Changes" : "Create Project"}
        </Btn>
      </div>
    </Modal>
  );
}

function AddEditModuleModal({ project, existing, onClose, onSave, T }) {
  const [f, sf] = useState(existing || { name:"", key:"", description:"" });
  const ch = k => e => sf(p => ({ ...p, [k]: e.target.value }));
  return (
    <Modal title={existing ? `Edit Module — ${project.name}` : `New Module — ${project.name}`} onClose={onClose} width={480} T={T}>
      <Field label="Module Name" required T={T}><TfInput T={T} value={f.name} onChange={ch("name")} placeholder="E.g. Authentication" /></Field>
      <Field label="Key" required T={T}><TfInput T={T} value={f.key} onChange={ch("key")} placeholder="E.g. AUTH" maxLength={8} /></Field>
      <Field label="Description" T={T}><TfTextarea T={T} value={f.description} onChange={ch("description")} /></Field>
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8, marginTop:8 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Cancel</Btn>
        <Btn variant="primary" T={T} onClick={() => { if (!f.name || !f.key) return; onSave({ ...(existing || {}), id:existing?.id || uid(), projectId:project.id, ...f, key:f.key.toUpperCase() }); }}>
          {existing ? "Save Changes" : "Create Module"}
        </Btn>
      </div>
    </Modal>
  );
}

function AddTCModal({ module, project, existing, tcCount, onClose, onSave, T }) {
  const nextNum = existing ? existing.num : tcCount + 1;
  const [f, sf] = useState(existing || { title:"", description:"", steps:"", expectedResult:"", priority:"medium", tags:"" });
  const ch = k => e => sf(p => ({ ...p, [k]: e.target.value }));
  const tcId = existing?.tcId || genTCId(project.key, module.key, nextNum);
  return (
    <Modal title={existing ? "Edit Test Case" : "New Test Case"} onClose={onClose} width={700} T={T}>
      <div style={{ background:T.blueSoft, borderRadius:8, padding:"8px 12px", marginBottom:16, fontSize:12, color:T.blueText, fontWeight:700, fontFamily:"monospace" }}>ID: {tcId}</div>
      <Field label="Title" required T={T}><TfInput T={T} value={f.title} onChange={ch("title")} placeholder="Short descriptive title" /></Field>
      <Field label="Description" T={T}><TfTextarea T={T} value={f.description} onChange={ch("description")} placeholder="What does this test verify?" /></Field>
      <Field label="Test Steps" required T={T}><TfTextarea T={T} minH={170} value={f.steps} onChange={ch("steps")} placeholder={"1. Navigate to...\n2. Click...\n3. Verify..."} /></Field>
      <Field label="Expected Result" required T={T}><TfTextarea T={T} value={f.expectedResult} onChange={ch("expectedResult")} placeholder="What should happen?" /></Field>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
        <Field label="Priority" T={T}>
          <TfSelect T={T} value={f.priority} onChange={ch("priority")}>
            <option value="high">High</option>
            <option value="medium">Medium</option>
            <option value="low">Low</option>
          </TfSelect>
        </Field>
        <Field label="Tags (comma separated)" T={T}><TfInput T={T} value={f.tags} onChange={ch("tags")} placeholder="smoke, regression" /></Field>
      </div>
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8, marginTop:8 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Cancel</Btn>
        <Btn variant="primary" T={T} onClick={() => { if (!f.title) return; onSave({ ...(existing || {}), id:existing?.id || uid(), moduleId:module.id, projectId:project.id, tcId, num:nextNum, ...f, status:"active" }); }}>
          {existing ? "Save Changes" : "Create Test Case"}
        </Btn>
      </div>
    </Modal>
  );
}

function ViewTCModal({ tc, onClose, onEdit, T }) {
  return (
    <Modal title={tc.tcId} onClose={onClose} width={640} T={T}>
      <div style={{ display:"flex", gap:8, marginBottom:12, flexWrap:"wrap" }}>
        <PriBadge p={tc.priority} T={T} />
        {tc.tags && tc.tags.split(",").map(t => <span key={t} style={{ fontSize:11, color:T.teal, background:T.tealSoft, padding:"2px 8px", borderRadius:20 }}>{t.trim()}</span>)}
      </div>
      <h3 style={{ margin:"0 0 6px", fontSize:16, color:T.text }}>{tc.title}</h3>
      {tc.description && <p style={{ color:T.textMuted, fontSize:13, margin:"0 0 16px" }}>{tc.description}</p>}
      <div style={{ background:T.surface2, borderRadius:8, padding:14, marginBottom:12 }}>
        <div style={{ fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:0.5, marginBottom:8 }}>Test Steps</div>
        <pre style={{ margin:0, fontSize:13, color:T.textSec, whiteSpace:"pre-wrap", fontFamily:"inherit", lineHeight:1.7 }}>{tc.steps}</pre>
      </div>
      <div style={{ background:T.greenSoft, borderRadius:8, padding:14 }}>
        <div style={{ fontSize:11, fontWeight:700, color:T.greenText, textTransform:"uppercase", letterSpacing:0.5, marginBottom:8 }}>Expected Result</div>
        <div style={{ fontSize:13, color:T.textSec }}>{tc.expectedResult}</div>
      </div>
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8, marginTop:16 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Close</Btn>
        <Btn onClick={onEdit} variant="primary" T={T}>Edit</Btn>
      </div>
    </Modal>
  );
}

// ─── Projects View ────────────────────────────────────────────────────────────
function ProjectsView({ state, dispatch, T, isMobile }) {
  const { projects, modules, testcases } = state;
  const [sel, setSel]         = useState(null);
  const [selMod, setSelMod]   = useState(null);
  const [projOpen, setProjOpen] = useState(true);
  const [modOpen, setModOpen]   = useState(true);
  const [searchProj, setSearchProj] = useState("");
  const [searchMod, setSearchMod]   = useState("");
  const [searchTC, setSearchTC]     = useState("");
  const [modal, setModal] = useState(null);
  const [mobileView, setMobileView] = useState("projects");

  const project  = projects.find(p => p.id === sel);
  const projMods = modules.filter(m => m.projectId === sel);
  const module   = projMods.find(m => m.id === selMod);
  const modTCs   = testcases.filter(tc => tc.moduleId === selMod);

  const filteredProjs = projects.filter(p => p.name.toLowerCase().includes(searchProj.toLowerCase()) || p.key.toLowerCase().includes(searchProj.toLowerCase()));
  const filteredMods  = projMods.filter(m => m.name.toLowerCase().includes(searchMod.toLowerCase()));
  const filteredTCs   = modTCs.filter(tc => tc.title.toLowerCase().includes(searchTC.toLowerCase()) || tc.tcId.toLowerCase().includes(searchTC.toLowerCase()) || tc.description.toLowerCase().includes(searchTC.toLowerCase()));

  const closeModal = () => setModal(null);

  function renderModals() {
    if (!modal) return null;
    const mTCCount = modTCs?.length || 0;
    if (modal.type === "addProject")    return <AddEditProjectModal onClose={closeModal} onSave={p => { dispatch({ type:"ADD_PROJECT", project:p }); closeModal(); }} T={T} />;
    if (modal.type === "editProject")   return <AddEditProjectModal existing={modal.data} onClose={closeModal} onSave={p => { dispatch({ type:"EDIT_PROJECT", project:p }); closeModal(); }} T={T} />;
    if (modal.type === "deleteProject") {
      const p = modal.data;
      const modCount = state.modules.filter(m => m.projectId === p.id).length;
      const tcCount  = state.testcases.filter(tc => tc.projectId === p.id).length;
      return <ConfirmModal title="Delete Project" message={`Deleting "${p.name}" will permanently remove ${modCount} module${modCount !== 1 ? "s" : ""} and ${tcCount} test case${tcCount !== 1 ? "s" : ""}. This cannot be undone.`} onConfirm={() => { dispatch({ type:"DELETE_PROJECT", id:p.id }); if (sel === p.id) { setSel(null); setSelMod(null); } }} onClose={closeModal} T={T} />;
    }
    if (modal.type === "addModule"    && project) return <AddEditModuleModal project={project} onClose={closeModal} onSave={m => { dispatch({ type:"ADD_MODULE", module:m }); closeModal(); }} T={T} />;
    if (modal.type === "editModule")  return <AddEditModuleModal project={project} existing={modal.data} onClose={closeModal} onSave={m => { dispatch({ type:"EDIT_MODULE", module:m }); closeModal(); }} T={T} />;
    if (modal.type === "deleteModule") {
      const m = modal.data;
      const tcCount = state.testcases.filter(tc => tc.moduleId === m.id).length;
      return <ConfirmModal title="Delete Module" message={`Deleting "${m.name}" will permanently remove ${tcCount} test case${tcCount !== 1 ? "s" : ""}. This cannot be undone.`} onConfirm={() => { dispatch({ type:"DELETE_MODULE", id:m.id }); if (selMod === m.id) setSelMod(null); }} onClose={closeModal} T={T} />;
    }
    if (modal.type === "addTC"  && module) return <AddTCModal module={module} project={project} tcCount={mTCCount} onClose={closeModal} onSave={tc => { dispatch({ type:"ADD_TC", tc }); closeModal(); }} T={T} />;
    if (modal.type === "editTC")           return <AddTCModal module={module} project={project} existing={modal.data} tcCount={mTCCount} onClose={closeModal} onSave={tc => { dispatch({ type:"EDIT_TC", tc }); closeModal(); }} T={T} />;
    if (modal.type === "viewTC")           return <ViewTCModal tc={modal.data} onClose={closeModal} onEdit={() => setModal({ type:"editTC", data:modal.data })} T={T} />;
    return null;
  }

  // ── Mobile layout ──
  if (isMobile) {
    return (
      <div style={{ display:"flex", flexDirection:"column", height:"100%" }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 14px", background:T.surface, borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
          {mobileView !== "projects" && <Btn onClick={() => { mobileView === "testcases" ? setMobileView("modules") : setMobileView("projects"); }} variant="ghost" T={T} style={{ padding:"4px 8px" }}>←</Btn>}
          <span style={{ fontSize:13, fontWeight:600, color:T.textMuted }}>{mobileView === "projects" ? "Projects" : mobileView === "modules" ? (project?.name || "Modules") : (module?.name || "Test Cases")}</span>
          <div style={{ flex:1 }} />
          {mobileView === "projects"   && <Btn onClick={() => setModal({ type:"addProject" })} variant="primary" T={T} style={{ padding:"5px 12px", fontSize:12 }}>+ New</Btn>}
          {mobileView === "modules"    && project && <Btn onClick={() => setModal({ type:"addModule" })} variant="primary" T={T} style={{ padding:"5px 12px", fontSize:12 }}>+ New</Btn>}
          {mobileView === "testcases"  && module  && <Btn onClick={() => setModal({ type:"addTC" })} variant="primary" T={T} style={{ padding:"5px 12px", fontSize:12 }}>+ New</Btn>}
        </div>
        <div style={{ padding:"8px 12px", background:T.surface, borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
          <SearchBar value={mobileView === "projects" ? searchProj : mobileView === "modules" ? searchMod : searchTC} onChange={mobileView === "projects" ? setSearchProj : mobileView === "modules" ? setSearchMod : setSearchTC} placeholder="Search…" T={T} />
        </div>
        <div style={{ flex:1, overflowY:"auto" }}>
          {mobileView === "projects" && filteredProjs.map(p => (
            <div key={p.id} style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"14px 16px", display:"flex", alignItems:"center", gap:10 }}>
              <div style={{ flex:1 }} onClick={() => { setSel(p.id); setSelMod(null); setMobileView("modules"); }}>
                <div style={{ fontSize:14, fontWeight:600, color:T.text }}>{p.name}</div>
                <div style={{ fontSize:12, color:T.textMuted, marginTop:2 }}>{p.key} · {testcases.filter(tc => tc.projectId === p.id).length} tests</div>
              </div>
              <Btn onClick={() => setModal({ type:"editProject", data:p })} variant="ghost" T={T} style={{ padding:"4px 8px", fontSize:12 }}>✎</Btn>
              <Btn onClick={() => setModal({ type:"deleteProject", data:p })} variant="ghost" T={T} style={{ padding:"4px 8px", fontSize:12, color:T.red }}>✕</Btn>
            </div>
          ))}
          {mobileView === "modules" && filteredMods.map(m => (
            <div key={m.id} style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"14px 16px", display:"flex", alignItems:"center", gap:10 }}>
              <div style={{ flex:1 }} onClick={() => { setSelMod(m.id); setMobileView("testcases"); }}>
                <div style={{ fontSize:14, fontWeight:600, color:T.text }}>{m.name}</div>
                <div style={{ fontSize:12, color:T.textMuted, marginTop:2 }}>{testcases.filter(tc => tc.moduleId === m.id).length} tests</div>
              </div>
              <Btn onClick={() => setModal({ type:"editModule", data:m })} variant="ghost" T={T} style={{ padding:"4px 8px", fontSize:12 }}>✎</Btn>
              <Btn onClick={() => setModal({ type:"deleteModule", data:m })} variant="ghost" T={T} style={{ padding:"4px 8px", fontSize:12, color:T.red }}>✕</Btn>
            </div>
          ))}
          {mobileView === "testcases" && (
            <div style={{ padding:12 }}>
              {filteredTCs.length === 0 && <div style={{ textAlign:"center", padding:32, color:T.textMuted }}>No test cases yet</div>}
              {filteredTCs.map(tc => (
                <div key={tc.id} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, padding:"12px 14px", marginBottom:8 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:4 }}>
                        <span style={{ fontSize:10, fontWeight:700, color:T.blue, background:T.blueSoft, padding:"2px 7px", borderRadius:5, fontFamily:"monospace" }}>{tc.tcId}</span>
                        <PriBadge p={tc.priority} T={T} />
                      </div>
                      <div style={{ fontSize:13, fontWeight:500, color:T.text }}>{tc.title}</div>
                    </div>
                    <div style={{ display:"flex", gap:2, flexShrink:0 }}>
                      <Btn onClick={() => setModal({ type:"viewTC", data:tc })} variant="ghost" T={T} style={{ padding:"3px 7px", fontSize:11 }}>👁</Btn>
                      <Btn onClick={() => setModal({ type:"editTC", data:tc })} variant="ghost" T={T} style={{ padding:"3px 7px", fontSize:11 }}>✎</Btn>
                      <Btn onClick={() => dispatch({ type:"DELETE_TC", id:tc.id })} variant="ghost" T={T} style={{ padding:"3px 7px", fontSize:11, color:T.red }}>✕</Btn>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
        {renderModals()}
      </div>
    );
  }

  // ── Desktop layout ──
  return (
    <div style={{ display:"flex", height:"100%", overflow:"hidden" }}>
      {/* Projects panel */}
      <div className="panel-slide" style={{ width:projOpen ? 220 : 44, minWidth:projOpen ? 220 : 44, flexShrink:0, borderRight:`1px solid ${T.border}`, background:T.surface2, overflowY:projOpen ? "auto" : "hidden", display:"flex", flexDirection:"column" }}>
        <div style={{ padding:"10px 12px", borderBottom:`1px solid ${T.border}`, display:"flex", justifyContent:"space-between", alignItems:"center", background:T.surface, flexShrink:0 }}>
          {projOpen && <span style={{ fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:0.8 }}>Projects</span>}
          <div style={{ display:"flex", gap:4 }}>
            {projOpen && <Btn onClick={() => setModal({ type:"addProject" })} variant="primary" T={T} style={{ padding:"2px 8px", fontSize:11 }}>+ New</Btn>}
            <Btn onClick={() => setProjOpen(v => !v)} variant="ghost" T={T} style={{ padding:"3px 6px", fontSize:12 }}>{projOpen ? "◀" : "▶"}</Btn>
          </div>
        </div>
        {projOpen && (
          <div style={{ padding:"8px 10px", borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
            <SearchBar value={searchProj} onChange={setSearchProj} placeholder="Search projects…" T={T} />
          </div>
        )}
        {projOpen && filteredProjs.map(p => (
          <div key={p.id} style={{ borderBottom:`1px solid ${T.border}`, background:sel === p.id ? T.navActive : "transparent", borderLeft:sel === p.id ? `3px solid ${T.blue}` : "3px solid transparent", transition:"all 0.12s" }}>
            <div style={{ padding:"9px 12px", cursor:"pointer", display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:4 }}>
              <div style={{ flex:1, minWidth:0 }} onClick={() => { setSel(p.id); setSelMod(null); }}>
                <div style={{ fontSize:13, fontWeight:600, color:sel === p.id ? T.blue : T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{p.name}</div>
                <div style={{ fontSize:11, color:T.textMuted, marginTop:1 }}>{p.key}</div>
              </div>
              <div style={{ display:"flex", gap:1, flexShrink:0 }}>
                <Btn onClick={() => setModal({ type:"editProject", data:p })} variant="ghost" T={T} style={{ padding:"2px 5px", fontSize:11, color:T.textMuted }}>✎</Btn>
                <Btn onClick={() => setModal({ type:"deleteProject", data:p })} variant="ghost" T={T} style={{ padding:"2px 5px", fontSize:11, color:T.red }}>✕</Btn>
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Modules panel */}
      <div className="panel-slide" style={{ width:modOpen ? 200 : 40, minWidth:modOpen ? 200 : 40, flexShrink:0, borderRight:`1px solid ${T.border}`, background:T.surface, overflowY:modOpen ? "auto" : "hidden", display:"flex", flexDirection:"column" }}>
        {project ? (
          <>
            <div style={{ padding:"10px 12px", borderBottom:`1px solid ${T.border}`, display:"flex", justifyContent:"space-between", alignItems:"center", flexShrink:0 }}>
              {modOpen && <span style={{ fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:0.8 }}>Modules</span>}
              <div style={{ display:"flex", gap:4 }}>
                {modOpen && <Btn onClick={() => setModal({ type:"addModule" })} variant="primary" T={T} style={{ padding:"2px 8px", fontSize:11 }}>+ New</Btn>}
                <Btn onClick={() => setModOpen(v => !v)} variant="ghost" T={T} style={{ padding:"3px 6px", fontSize:12 }}>{modOpen ? "◀" : "▶"}</Btn>
              </div>
            </div>
            {modOpen && (
              <div style={{ padding:"8px 10px", borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
                <SearchBar value={searchMod} onChange={setSearchMod} placeholder="Search modules…" T={T} />
              </div>
            )}
            {modOpen && filteredMods.length === 0 && <div style={{ padding:14, color:T.textMuted, fontSize:12 }}>No modules yet</div>}
            {modOpen && filteredMods.map(m => (
              <div key={m.id} style={{ borderBottom:`1px solid ${T.border}`, background:selMod === m.id ? T.navActive : "transparent", borderLeft:selMod === m.id ? `3px solid ${T.blue}` : "3px solid transparent", transition:"all 0.12s" }}>
                <div style={{ padding:"9px 12px", display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:4 }}>
                  <div style={{ flex:1, minWidth:0, cursor:"pointer" }} onClick={() => setSelMod(m.id)}>
                    <div style={{ fontSize:13, fontWeight:500, color:selMod === m.id ? T.blue : T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{m.name}</div>
                    <div style={{ fontSize:11, color:T.textMuted, marginTop:1 }}>{testcases.filter(tc => tc.moduleId === m.id).length} tests</div>
                  </div>
                  <div style={{ display:"flex", gap:1, flexShrink:0 }}>
                    <Btn onClick={() => setModal({ type:"editModule", data:m })} variant="ghost" T={T} style={{ padding:"2px 5px", fontSize:11, color:T.textMuted }}>✎</Btn>
                    <Btn onClick={() => setModal({ type:"deleteModule", data:m })} variant="ghost" T={T} style={{ padding:"2px 5px", fontSize:11, color:T.red }}>✕</Btn>
                  </div>
                </div>
              </div>
            ))}
          </>
        ) : (
          <div style={{ padding:14, color:T.textMuted, fontSize:12 }}>{modOpen ? "Select a project" : ""}</div>
        )}
      </div>

      {/* Test Cases panel */}
      <div style={{ flex:1, overflowY:"auto", display:"flex", flexDirection:"column" }}>
        {module ? (
          <>
            <div style={{ padding:"12px 18px", borderBottom:`1px solid ${T.border}`, display:"flex", justifyContent:"space-between", alignItems:"center", background:T.surface, position:"sticky", top:0, zIndex:1, flexShrink:0 }}>
              <div>
                <div style={{ fontSize:15, fontWeight:700, color:T.text }}>{module.name}</div>
                <div style={{ fontSize:12, color:T.textMuted }}>{module.description || project?.name}</div>
              </div>
              <Btn onClick={() => setModal({ type:"addTC" })} variant="primary" T={T}>+ New Test Case</Btn>
            </div>
            <div style={{ padding:"10px 14px", borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
              <SearchBar value={searchTC} onChange={setSearchTC} placeholder="Search test cases…" T={T} />
            </div>
            <div style={{ padding:14 }}>
              {filteredTCs.length === 0 && <div style={{ textAlign:"center", padding:36, color:T.textMuted }}><div style={{ fontSize:28, marginBottom:8 }}>📝</div>No test cases yet</div>}
              {filteredTCs.map(tc => (
                <div key={tc.id} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, padding:"12px 14px", marginBottom:8, display:"flex", alignItems:"flex-start", gap:10 }}>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ display:"flex", gap:7, flexWrap:"wrap", marginBottom:5 }}>
                      <span style={{ fontSize:11, fontWeight:700, color:T.blue, background:T.blueSoft, padding:"2px 8px", borderRadius:5, fontFamily:"monospace" }}>{tc.tcId}</span>
                      <PriBadge p={tc.priority} T={T} />
                    </div>
                    <div style={{ fontSize:14, fontWeight:500, color:T.text }}>{tc.title}</div>
                    <div style={{ fontSize:12, color:T.textMuted, marginTop:2, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{tc.description}</div>
                  </div>
                  <div style={{ display:"flex", gap:3, flexShrink:0 }}>
                    <Btn onClick={() => setModal({ type:"viewTC", data:tc })} variant="outline" T={T} style={{ padding:"4px 10px", fontSize:12 }}>View</Btn>
                    <Btn onClick={() => setModal({ type:"editTC", data:tc })} variant="outline" T={T} style={{ padding:"4px 10px", fontSize:12 }}>Edit</Btn>
                    <Btn onClick={() => dispatch({ type:"DELETE_TC", id:tc.id })} variant="ghost" T={T} style={{ padding:"4px 8px", fontSize:12, color:T.red }}>✕</Btn>
                  </div>
                </div>
              ))}
            </div>
          </>
        ) : project ? (
          <div style={{ padding:40, color:T.textMuted, fontSize:14, textAlign:"center" }}><div style={{ fontSize:32, marginBottom:8 }}>📂</div>Select a module to view test cases</div>
        ) : (
          <div style={{ padding:40, color:T.textMuted, fontSize:14, textAlign:"center" }}><div style={{ fontSize:32, marginBottom:8 }}>📁</div>Select a project to get started</div>
        )}
      </div>
      {renderModals()}
    </div>
  );
}

// ─── Suites View ──────────────────────────────────────────────────────────────
function AddSuiteModal({ existing, state, onClose, onSave, T }) {
  const { projects, modules, testcases } = state;
  const [f, sf]           = useState({ name:existing?.name || "", description:existing?.description || "" });
  const [inclusions, setInc] = useState(existing?.inclusions || []);
  const [addType, setAddType] = useState("project");
  const [addId, setAddId]     = useState("");

  const getLabel = (type, id) => {
    if (type === "project") return projects.find(p => p.id === id)?.name || id;
    if (type === "module")  return modules.find(m => m.id === id)?.name || id;
    return testcases.find(tc => tc.id === id)?.tcId || id;
  };
  const opts = { project:projects, module:modules, tc:testcases };
  const addInc = () => {
    if (!addId) return;
    if (inclusions.find(i => i.type === addType && i.id === addId)) return;
    setInc(p => [...p, { type:addType, id:addId, label:getLabel(addType, addId) }]);
    setAddId("");
  };

  return (
    <Modal title={existing ? "Edit Suite" : "New Test Suite"} onClose={onClose} width={580} T={T}>
      <Field label="Suite Name" required T={T}><TfInput T={T} value={f.name} onChange={e => sf(p => ({ ...p, name:e.target.value }))} placeholder="E.g. Smoke Tests" /></Field>
      <Field label="Description" T={T}><TfTextarea T={T} value={f.description} onChange={e => sf(p => ({ ...p, description:e.target.value }))} /></Field>
      <div style={{ fontSize:13, fontWeight:600, color:T.textSec, marginBottom:10 }}>Add Test Cases</div>
      <div style={{ display:"flex", gap:8, marginBottom:12 }}>
        <TfSelect T={T} value={addType} onChange={e => { setAddType(e.target.value); setAddId(""); }} style={{ width:120 }}>
          <option value="project">Project</option>
          <option value="module">Module</option>
          <option value="tc">Test Case</option>
        </TfSelect>
        <TfSelect T={T} value={addId} onChange={e => setAddId(e.target.value)} style={{ flex:1 }}>
          <option value="">Select…</option>
          {opts[addType].map(item => <option key={item.id} value={item.id}>{item.name || item.title || item.tcId}</option>)}
        </TfSelect>
        <Btn variant="primary" onClick={addInc} T={T} style={{ padding:"6px 14px" }}>Add</Btn>
      </div>
      {inclusions.length > 0 && (
        <div style={{ background:T.surface2, borderRadius:8, padding:10, marginBottom:16 }}>
          {inclusions.map((inc, i) => (
            <div key={i} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"5px 0", borderBottom:i < inclusions.length - 1 ? `1px solid ${T.border}` : "none" }}>
              <span style={{ fontSize:12, color:T.textSec }}><span style={{ color:T.textMuted }}>{inc.type}: </span>{inc.label}</span>
              <Btn variant="ghost" T={T} style={{ padding:"2px 6px", fontSize:11, color:T.red }} onClick={() => setInc(p => p.filter((_, j) => j !== i))}>✕</Btn>
            </div>
          ))}
        </div>
      )}
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Cancel</Btn>
        <Btn variant="primary" T={T} onClick={() => { if (!f.name) return; onSave({ id:existing?.id || uid(), ...f, inclusions, createdAt:existing?.createdAt || new Date().toISOString() }); }}>
          {existing ? "Save Changes" : "Create Suite"}
        </Btn>
      </div>
    </Modal>
  );
}

function SuitesView({ state, dispatch, T }) {
  const [showAdd, setShowAdd]   = useState(false);
  const [editSuite, setEditSuite] = useState(null);
  const [search, setSearch]     = useState("");
  const [confirm, setConfirm]   = useState(null);
  const filtered = state.suites.filter(s => s.name.toLowerCase().includes(search.toLowerCase()));
  return (
    <div style={{ padding:20, maxWidth:900, margin:"0 auto", overflowY:"auto", height:"100%" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <div>
          <h2 style={{ margin:0, fontSize:18, fontWeight:700, color:T.text }}>Test Suites</h2>
          <p style={{ margin:"3px 0 0", color:T.textMuted, fontSize:13 }}>Group test cases into reusable suites</p>
        </div>
        <Btn variant="primary" T={T} onClick={() => setShowAdd(true)}>+ New Suite</Btn>
      </div>
      <div style={{ marginBottom:14 }}><SearchBar value={search} onChange={setSearch} placeholder="Search suites…" T={T} /></div>
      {state.suites.length === 0 && (
        <div style={{ textAlign:"center", padding:"50px 40px", background:T.surface, borderRadius:14, border:`1px dashed ${T.border}` }}>
          <div style={{ fontSize:36, marginBottom:10 }}>📋</div>
          <div style={{ fontSize:14, fontWeight:500, color:T.textSec, marginBottom:4 }}>No test suites yet</div>
          <div style={{ fontSize:13, color:T.textMuted }}>Create a suite to group test cases for a test run</div>
        </div>
      )}
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {filtered.map(suite => (
          <div key={suite.id} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:12, padding:"14px 18px", display:"flex", alignItems:"center", gap:14 }}>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:14, fontWeight:600, color:T.text, marginBottom:3 }}>{suite.name}</div>
              {suite.description && <div style={{ fontSize:12, color:T.textMuted, marginBottom:6 }}>{suite.description}</div>}
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                <span style={{ fontSize:11, color:T.blue, background:T.blueSoft, padding:"2px 8px", borderRadius:20 }}>{resolveSuiteTCs(suite, state).length} test cases</span>
                {suite.inclusions.map((inc, i) => <span key={i} style={{ fontSize:11, color:T.textMuted, background:T.gray100, padding:"2px 8px", borderRadius:20 }}>{inc.type}: {inc.label}</span>)}
              </div>
            </div>
            <div style={{ display:"flex", gap:4 }}>
              <Btn variant="outline" T={T} style={{ fontSize:12, padding:"5px 11px" }} onClick={() => setEditSuite(suite)}>Edit</Btn>
              <Btn variant="ghost" T={T} style={{ fontSize:12, color:T.red }} onClick={() => setConfirm(suite)}>✕</Btn>
            </div>
          </div>
        ))}
      </div>
      {(showAdd || editSuite) && <AddSuiteModal existing={editSuite} state={state} onClose={() => { setShowAdd(false); setEditSuite(null); }} onSave={s => { dispatch({ type:editSuite ? "EDIT_SUITE" : "ADD_SUITE", suite:s }); setShowAdd(false); setEditSuite(null); }} T={T} />}
      {confirm && <ConfirmModal title="Delete Suite" message={`Delete "${confirm.name}"? This won't affect the test cases themselves.`} onConfirm={() => { dispatch({ type:"DELETE_SUITE", id:confirm.id }); setConfirm(null); }} onClose={() => setConfirm(null)} T={T} />}
    </div>
  );
}

// ─── Runs View ────────────────────────────────────────────────────────────────
function AddRunModal({ state, onClose, onSave, T }) {
  const { suites } = state;
  const [name, setName]           = useState("");
  const [desc, setDesc]           = useState("");
  const [selSuites, setSelSuites] = useState([]);
  const toggle = id => setSelSuites(p => p.includes(id) ? p.filter(x => x !== id) : [...p, id]);
  const tcCount = useMemo(() => {
    const set = new Set();
    selSuites.forEach(sid => { const suite = suites.find(s => s.id === sid); if (suite) resolveSuiteTCs(suite, state).forEach(tc => set.add(tc.id)); });
    return set.size;
  }, [selSuites, state, suites]);
  return (
    <Modal title="New Test Run" onClose={onClose} width={520} T={T}>
      <Field label="Run Name" required T={T}><TfInput T={T} value={name} onChange={e => setName(e.target.value)} placeholder="E.g. Sprint 12 Regression" /></Field>
      <Field label="Description" T={T}><TfTextarea T={T} value={desc} onChange={e => setDesc(e.target.value)} /></Field>
      <Field label="Select Suites" required T={T}>
        <div style={{ border:`1px solid ${T.border}`, borderRadius:8, overflow:"hidden" }}>
          {suites.length === 0 && <div style={{ padding:"12px 14px", fontSize:13, color:T.textMuted }}>No suites available</div>}
          {suites.map(s => {
            const checked = selSuites.includes(s.id);
            return (
              <label key={s.id} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 14px", cursor:"pointer", background:checked ? T.navActive : T.surface, borderBottom:`1px solid ${T.border}` }}>
                <input type="checkbox" checked={checked} onChange={() => toggle(s.id)} />
                <div>
                  <div style={{ fontSize:13, fontWeight:500, color:T.text }}>{s.name}</div>
                  <div style={{ fontSize:11, color:T.textMuted }}>{resolveSuiteTCs(s, state).length} test cases</div>
                </div>
              </label>
            );
          })}
        </div>
      </Field>
      {tcCount > 0 && <div style={{ fontSize:13, color:T.blue, marginBottom:12 }}>→ {tcCount} unique test cases will be included</div>}
      <div style={{ display:"flex", justifyContent:"flex-end", gap:8 }}>
        <Btn onClick={onClose} variant="outline" T={T}>Cancel</Btn>
        <Btn variant="primary" T={T} onClick={() => {
          if (!name || selSuites.length === 0) return;
          const tcSet = new Set();
          selSuites.forEach(sid => { const suite = suites.find(s => s.id === sid); if (suite) resolveSuiteTCs(suite, state).forEach(tc => tcSet.add(tc.id)); });
          const cases = [...tcSet].map(id => ({ tcId:id, status:"pending", comment:"", executedAt:null }));
          onSave({ id:uid(), name, description:desc, suiteIds:selSuites, cases, status:"active", createdAt:new Date().toISOString(), completedAt:null });
        }}>Start Run</Btn>
      </div>
    </Modal>
  );
}

function CaseExecutor({ runCase, tc, state, onUpdate, disabled, T }) {
  if (!tc) return null;
  const mod  = state.modules.find(m => m.id === tc.moduleId);
  const proj = state.projects.find(p => p.id === tc.projectId);
  return (
    <div style={{ padding:22 }}>
      <div style={{ marginBottom:14 }}>
        <div style={{ fontSize:11, color:T.textMuted, marginBottom:4 }}>{proj?.name} › {mod?.name}</div>
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:4 }}>
          <span style={{ fontFamily:"monospace", fontSize:12, color:T.blue, background:T.blueSoft, padding:"2px 8px", borderRadius:5 }}>{tc.tcId}</span>
          <PriBadge p={tc.priority} T={T} />
        </div>
        <h3 style={{ margin:"8px 0 4px", fontSize:16, fontWeight:700, color:T.text }}>{tc.title}</h3>
        {tc.description && <p style={{ margin:"0 0 14px", color:T.textMuted, fontSize:13 }}>{tc.description}</p>}
      </div>
      <div style={{ background:T.surface2, borderRadius:8, padding:14, marginBottom:12 }}>
        <div style={{ fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:0.5, marginBottom:8 }}>Test Steps</div>
        <pre style={{ margin:0, fontSize:13, color:T.textSec, whiteSpace:"pre-wrap", fontFamily:"inherit", lineHeight:1.7 }}>{tc.steps}</pre>
      </div>
      <div style={{ background:T.greenSoft, borderRadius:8, padding:14, marginBottom:18 }}>
        <div style={{ fontSize:11, fontWeight:700, color:T.greenText, textTransform:"uppercase", letterSpacing:0.5, marginBottom:8 }}>Expected Result</div>
        <div style={{ fontSize:13, color:T.textSec }}>{tc.expectedResult}</div>
      </div>
      {!disabled && (
        <>
          <div style={{ fontSize:13, fontWeight:600, color:T.textSec, marginBottom:10 }}>Update Result</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:7, marginBottom:14 }}>
            {Object.entries(STATUS_META).filter(([k]) => k !== "in-progress").map(([k, m]) => {
              const active = runCase.status === k;
              const color  = T[m.colorKey] || T.textMuted;
              const bg     = T[m.colorKey + "Soft"] || T.gray100;
              return <Btn key={k} onClick={() => onUpdate({ status:k, executedAt:new Date().toISOString() })} variant="outline" T={T} style={{ background:active ? bg : T.surface, borderColor:active ? color : T.border, color:active ? color : T.textSec, fontWeight:active ? 700 : 500, padding:"7px 14px" }}>{m.icon} {m.label}</Btn>;
            })}
          </div>
          <Field label="Comment / Notes" T={T}>
            <TfTextarea T={T} minH={70} value={runCase.comment} onChange={e => onUpdate({ comment:e.target.value })} placeholder="Optional: describe what you observed…" />
          </Field>
        </>
      )}
      {disabled && (
        <div style={{ background:T.surface2, borderRadius:8, padding:14 }}>
          <div style={{ fontSize:12, color:T.textMuted, marginBottom:6 }}>Result</div>
          <StatusBadge status={runCase.status} T={T} />
          {runCase.comment && <div style={{ fontSize:13, color:T.textSec, marginTop:8 }}>{runCase.comment}</div>}
        </div>
      )}
    </div>
  );
}

function RunExecutionView({ run, state, dispatch, onBack, T, isMobile }) {
  const [filter, setFilter] = useState("all");
  const [sel, setSel]       = useState(null);
  const [showList, setShowList] = useState(true);

  const getTC   = id => state.testcases.find(t => t.id === id);
  const cases   = run.cases;
  const filtered = filter === "all" ? cases : cases.filter(c => c.status === filter);
  const stats   = useMemo(() => {
    const total = cases.length, counts = {};
    Object.keys(STATUS_META).forEach(k => { counts[k] = cases.filter(c => c.status === k).length; });
    return { total, ...counts, done: total - counts.pending };
  }, [cases]);

  const updateCase = (tcId, update) => dispatch({ type:"UPDATE_RUN_CASE", runId:run.id, tcId, update });
  const completeRun = () => dispatch({ type:"COMPLETE_RUN", runId:run.id });
  const allDone = stats.pending === 0;
  const pct     = stats.total ? Math.round((stats.pass / stats.total) * 100) : 0;

  const header = (
    <div style={{ padding:isMobile ? "10px 14px" : "12px 18px", borderBottom:`1px solid ${T.border}`, background:T.surface, display:"flex", alignItems:"center", gap:10, flexShrink:0, flexWrap:"wrap" }}>
      <Btn variant="ghost" onClick={onBack} T={T} style={{ padding:"4px 8px" }}>← Back</Btn>
      <div style={{ flex:1, minWidth:0 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
          <span style={{ fontSize:14, fontWeight:700, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:180 }}>{run.name}</span>
          <RunStatusBadge status={run.status} T={T} />
        </div>
      </div>
      <div style={{ display:"flex", gap:12, fontSize:12 }}>
        <span style={{ color:T.green }}>✓ {stats.pass}</span>
        <span style={{ color:T.red }}>✗ {stats.fail}</span>
        <span style={{ color:T.textMuted }}>⋯ {stats.pending}</span>
        <span style={{ fontWeight:700, color:T.blue }}>{pct}%</span>
      </div>
      {run.status !== "completed" && allDone  && <Btn variant="success" onClick={completeRun} T={T}>✓ Complete</Btn>}
      {run.status !== "completed" && !allDone && <span style={{ fontSize:11, color:T.textMuted }}>{stats.done}/{stats.total}</span>}
    </div>
  );

  const progress = <div style={{ height:4, background:T.surface2, flexShrink:0 }}><div style={{ height:"100%", background:T.green, width:pct + "%", transition:"width 0.35s" }} /></div>;

  const caseList = (
    <div style={{ overflowY:"auto", flex:1, display:"flex", flexDirection:"column" }}>
      <div style={{ padding:"8px 10px", borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
        <TfSelect T={T} value={filter} onChange={e => setFilter(e.target.value)} style={{ fontSize:12 }}>
          <option value="all">All ({stats.total})</option>
          <option value="pending">Pending ({stats.pending})</option>
          <option value="pass">Pass ({stats.pass})</option>
          <option value="fail">Fail ({stats.fail})</option>
          <option value="blocked">Blocked ({stats.blocked})</option>
          <option value="skipped">Skipped ({stats.skipped})</option>
        </TfSelect>
      </div>
      {filtered.map(c => {
        const tc  = getTC(c.tcId); if (!tc) return null;
        const mod = state.modules.find(m => m.id === tc.moduleId);
        return (
          <div key={c.tcId} onClick={() => { setSel(c.tcId); setShowList(false); }}
            style={{ padding:"10px 13px", cursor:"pointer", borderBottom:`1px solid ${T.border}`, background:sel === c.tcId ? T.navActive : T.surface, borderLeft:sel === c.tcId ? `3px solid ${T.blue}` : "3px solid transparent" }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:6 }}>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:10, color:T.textMuted, fontFamily:"monospace", marginBottom:2 }}>{tc.tcId}</div>
                <div style={{ fontSize:12, fontWeight:500, color:T.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{tc.title}</div>
                {!isMobile && <div style={{ fontSize:11, color:T.textMuted, marginTop:2 }}>{mod?.name}</div>}
              </div>
              <StatusBadge status={c.status} T={T} small />
            </div>
          </div>
        );
      })}
    </div>
  );

  const caseDetail = (
    <div style={{ flex:1, overflowY:"auto" }}>
      {sel
        ? <CaseExecutor runCase={cases.find(c => c.tcId === sel)} tc={getTC(sel)} state={state} disabled={run.status === "completed"} onUpdate={u => updateCase(sel, u)} T={T} />
        : <div style={{ display:"flex", alignItems:"center", justifyContent:"center", height:"100%", color:T.textMuted, fontSize:14 }}>Select a test case to execute</div>
      }
    </div>
  );

  if (isMobile) {
    return (
      <div style={{ display:"flex", flexDirection:"column", height:"100%" }}>
        {header}{progress}
        {sel && (
          <div style={{ display:"flex", background:T.surface, borderBottom:`1px solid ${T.border}`, flexShrink:0 }}>
            <button onClick={() => setSel(null)} style={{ flex:1, padding:8, border:"none", background:showList ? T.navActive : T.surface, color:showList ? T.blue : T.textMuted, fontFamily:"inherit", fontSize:12, fontWeight:600, cursor:"pointer" }}>Case List</button>
            <button onClick={() => setShowList(false)} style={{ flex:1, padding:8, border:"none", background:!showList ? T.navActive : T.surface, color:!showList ? T.blue : T.textMuted, fontFamily:"inherit", fontSize:12, fontWeight:600, cursor:"pointer" }}>Details</button>
          </div>
        )}
        <div style={{ flex:1, overflow:"hidden", display:"flex", flexDirection:"column" }}>
          {(!sel || showList) && caseList}
          {sel && !showList && caseDetail}
        </div>
      </div>
    );
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", height:"100%" }}>
      {header}{progress}
      <div style={{ display:"flex", flex:1, overflow:"hidden" }}>
        <div style={{ width:300, flexShrink:0, borderRight:`1px solid ${T.border}`, display:"flex", flexDirection:"column" }}>{caseList}</div>
        {caseDetail}
      </div>
    </div>
  );
}

function RunsView({ state, dispatch, T, isMobile }) {
  const { runs, suites } = state;
  const [showAdd, setShowAdd]   = useState(false);
  const [activeRun, setActiveRun] = useState(null);
  const [search, setSearch]     = useState("");
  const [confirm, setConfirm]   = useState(null);
  const run = runs.find(r => r.id === activeRun);

  if (run) return <RunExecutionView run={run} state={state} dispatch={dispatch} onBack={() => setActiveRun(null)} T={T} isMobile={isMobile} />;

  const filtered = runs.filter(r => r.name.toLowerCase().includes(search.toLowerCase()));
  return (
    <div style={{ padding:20, maxWidth:900, margin:"0 auto", overflowY:"auto", height:"100%" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
        <div>
          <h2 style={{ margin:0, fontSize:18, fontWeight:700, color:T.text }}>Test Runs</h2>
          <p style={{ margin:"3px 0 0", color:T.textMuted, fontSize:13 }}>Execute and track test results</p>
        </div>
        <Btn variant="primary" T={T} onClick={() => setShowAdd(true)} disabled={suites.length === 0}>+ New Run</Btn>
      </div>
      <div style={{ marginBottom:14 }}><SearchBar value={search} onChange={setSearch} placeholder="Search runs…" T={T} /></div>
      {suites.length === 0 && <div style={{ background:T.amberSoft, borderRadius:10, padding:"10px 14px", marginBottom:16, fontSize:13, color:T.amber }}>⚠ Create at least one test suite before starting a run.</div>}
      {runs.length === 0 && <div style={{ textAlign:"center", padding:"50px 40px", background:T.surface, borderRadius:14, border:`1px dashed ${T.border}` }}><div style={{ fontSize:36, marginBottom:10 }}>▶</div><div style={{ fontSize:14, fontWeight:500, color:T.textSec }}>No test runs yet</div></div>}
      <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
        {filtered.map(r => {
          const total  = r.cases.length;
          const passed = r.cases.filter(c => c.status === "pass").length;
          const failed = r.cases.filter(c => c.status === "fail").length;
          const pending = r.cases.filter(c => c.status === "pending").length;
          const pct    = total ? Math.round((passed / total) * 100) : 0;
          return (
            <div key={r.id} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:12, padding:"14px 18px", display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
              <div style={{ flex:1, minWidth:200 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:4, flexWrap:"wrap" }}>
                  <span style={{ fontSize:14, fontWeight:600, color:T.text }}>{r.name}</span>
                  <RunStatusBadge status={r.status} T={T} />
                </div>
                <div style={{ fontSize:12, color:T.textMuted, marginBottom:6 }}>{r.suiteIds.length} suite{r.suiteIds.length > 1 ? "s" : ""} · {total} test cases</div>
                <div style={{ display:"flex", gap:12, fontSize:12, flexWrap:"wrap" }}>
                  <span style={{ color:T.green }}>✓ {passed} passed</span>
                  <span style={{ color:T.red }}>✗ {failed} failed</span>
                  <span style={{ color:T.textMuted }}>⋯ {pending} pending</span>
                  <span style={{ fontWeight:700, color:T.blue }}>{pct}%</span>
                </div>
              </div>
              <div style={{ display:"flex", gap:5 }}>
                {r.status !== "completed" && <Btn variant="primary" T={T} style={{ fontSize:12 }} onClick={() => setActiveRun(r.id)}>Execute</Btn>}
                {r.status === "completed" && <Btn variant="outline" T={T} style={{ fontSize:12 }} onClick={() => setActiveRun(r.id)}>View</Btn>}
                <Btn variant="ghost" T={T} style={{ fontSize:12, color:T.red }} onClick={() => setConfirm(r)}>✕</Btn>
              </div>
            </div>
          );
        })}
      </div>
      {showAdd && <AddRunModal state={state} onClose={() => setShowAdd(false)} onSave={r => { dispatch({ type:"ADD_RUN", run:r }); setShowAdd(false); setActiveRun(r.id); }} T={T} />}
      {confirm && <ConfirmModal title="Delete Run" message={`Delete test run "${confirm.name}"? All execution data will be lost.`} onConfirm={() => { dispatch({ type:"DELETE_RUN", id:confirm.id }); setConfirm(null); }} onClose={() => setConfirm(null)} T={T} />}
    </div>
  );
}

// ─── Reports View ─────────────────────────────────────────────────────────────
function DonutChart({ passed, failed, other, total, T }) {
  if (!total) return null;
  const cx = 80, cy = 80, r = 60, sw = 20;
  let cum = -90;
  const segs = [{ v:passed, color:T.green }, { v:failed, color:T.red }, { v:other, color:T.border }].filter(s => s.v > 0);
  const arcs = segs.map(s => { const a = (s.v / total) * 360, start = cum; cum += a; return { ...s, start, angle:a }; });
  const pxy  = (a, rr) => { const rad = a * Math.PI / 180; return { x: cx + rr * Math.cos(rad), y: cy + rr * Math.sin(rad) }; };
  return (
    <svg width={160} height={160} viewBox="0 0 160 160" style={{ display:"block", margin:"0 auto" }}>
      {arcs.map((arc, i) => {
        if (!arc.angle) return null;
        const s = pxy(arc.start, r), e = pxy(arc.start + arc.angle, r), large = arc.angle > 180 ? 1 : 0;
        return <path key={i} d={`M ${s.x} ${s.y} A ${r} ${r} 0 ${large} 1 ${e.x} ${e.y}`} fill="none" stroke={arc.color} strokeWidth={sw} />;
      })}
      <text x={cx} y={cy - 5} textAnchor="middle" fontSize={22} fontWeight={800} fill={T.text}>{Math.round((passed / total) * 100)}%</text>
      <text x={cx} y={cy + 14} textAnchor="middle" fontSize={10} fill={T.textMuted}>pass rate</text>
    </svg>
  );
}

function ReportContent({ run, state, T }) {
  const cases = run.cases, total = cases.length;
  const byStatus = {};
  Object.keys(STATUS_META).forEach(k => { byStatus[k] = cases.filter(c => c.status === k).length; });
  const passed = byStatus.pass || 0, failed = byStatus.fail || 0;
  const passRate = total ? Math.round((passed / total) * 100) : 0;

  const byModule = {};
  cases.forEach(c => {
    const tc  = state.testcases.find(t => t.id === c.tcId); if (!tc) return;
    const mod = state.modules.find(m => m.id === tc.moduleId);
    const key = mod?.name || "Unknown";
    if (!byModule[key]) byModule[key] = { pass:0, fail:0, total:0 };
    byModule[key].total++;
    if (c.status === "pass") byModule[key].pass++;
    else if (c.status === "fail") byModule[key].fail++;
  });

  const byPriority = {};
  ["high", "medium", "low"].forEach(p => {
    const rel = cases.filter(c => state.testcases.find(t => t.id === c.tcId)?.priority === p);
    byPriority[p] = { total:rel.length, pass:rel.filter(c => c.status === "pass").length, fail:rel.filter(c => c.status === "fail").length };
  });

  const rc = passRate >= 80 ? T.green : passRate >= 60 ? T.amber : T.red;

  return (
    <div>
      {/* Summary card */}
      <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:14, padding:20, marginBottom:14 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:12 }}>
          <div>
            <h3 style={{ margin:"0 0 4px", fontSize:17, fontWeight:700, color:T.text }}>{run.name}</h3>
            <div style={{ fontSize:12, color:T.textMuted }}>Completed {run.completedAt ? new Date(run.completedAt).toLocaleString() : "—"}</div>
          </div>
          <div style={{ textAlign:"right" }}>
            <div style={{ fontSize:38, fontWeight:800, color:rc }}>{passRate}%</div>
            <div style={{ fontSize:11, color:T.textMuted }}>Pass Rate</div>
          </div>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(80px,1fr))", gap:8, marginTop:16 }}>
          {[{ label:"Total", val:total, color:T.text }, { label:"Passed", val:byStatus.pass||0, color:T.green }, { label:"Failed", val:byStatus.fail||0, color:T.red }, { label:"Blocked", val:byStatus.blocked||0, color:T.amber }, { label:"Skipped", val:byStatus.skipped||0, color:T.purple }, { label:"Pending", val:byStatus.pending||0, color:T.textMuted }].map(m => (
            <div key={m.label} style={{ background:T.surface2, borderRadius:10, padding:"10px 12px", textAlign:"center" }}>
              <div style={{ fontSize:20, fontWeight:800, color:m.color }}>{m.val}</div>
              <div style={{ fontSize:11, color:T.textMuted, marginTop:2 }}>{m.label}</div>
            </div>
          ))}
        </div>
        <div style={{ marginTop:14, height:8, background:T.surface2, borderRadius:8, overflow:"hidden", display:"flex" }}>
          {passed > 0 && <div style={{ width:`${(passed/total)*100}%`, background:T.green }} />}
          {failed > 0 && <div style={{ width:`${(failed/total)*100}%`, background:T.red }} />}
          {(byStatus.blocked||0) > 0 && <div style={{ width:`${((byStatus.blocked||0)/total)*100}%`, background:T.amber }} />}
          {(byStatus.skipped||0) > 0 && <div style={{ width:`${((byStatus.skipped||0)/total)*100}%`, background:T.purple }} />}
        </div>
      </div>

      {/* Charts */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))", gap:14, marginBottom:14 }}>
        <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:14, padding:18 }}>
          <h4 style={{ margin:"0 0 14px", fontSize:14, fontWeight:600, color:T.textSec }}>Results by Module</h4>
          {Object.keys(byModule).length === 0 && <div style={{ color:T.textMuted, fontSize:13 }}>No data</div>}
          {Object.entries(byModule).map(([name, data]) => {
            const p = data.total ? Math.round((data.pass / data.total) * 100) : 0;
            const c = p >= 80 ? T.green : p >= 60 ? T.amber : T.red;
            return (
              <div key={name} style={{ marginBottom:12 }}>
                <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:T.textSec, marginBottom:4 }}><span>{name}</span><span style={{ fontWeight:600, color:c }}>{p}%</span></div>
                <div style={{ height:6, background:T.surface2, borderRadius:4, overflow:"hidden" }}><div style={{ width:p + "%", background:c, height:"100%", borderRadius:4, transition:"width 0.4s" }} /></div>
                <div style={{ fontSize:11, color:T.textMuted, marginTop:2 }}>✓{data.pass} ✗{data.fail} of {data.total}</div>
              </div>
            );
          })}
        </div>
        <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:14, padding:18 }}>
          <h4 style={{ margin:"0 0 14px", fontSize:14, fontWeight:600, color:T.textSec }}>Results by Priority</h4>
          {[["high","High",T.red],["medium","Medium",T.amber],["low","Low",T.green]].map(([key, label, color]) => {
            const data = byPriority[key];
            const p    = data.total ? Math.round((data.pass / data.total) * 100) : 0;
            return (
              <div key={key} style={{ marginBottom:12 }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:6 }}><span style={{ width:8, height:8, borderRadius:2, background:color, display:"inline-block" }} /><span style={{ fontSize:12, color:T.textSec }}>{label}</span></div>
                  <span style={{ fontSize:12, color:T.textMuted }}>{data.total} tests</span>
                </div>
                <div style={{ height:18, background:T.surface2, borderRadius:4, overflow:"hidden", display:"flex" }}>
                  {data.pass > 0 && <div style={{ width:`${(data.pass/data.total)*100}%`, background:T.green, display:"flex", alignItems:"center", justifyContent:"center" }}><span style={{ fontSize:9, color:"white", fontWeight:700 }}>{data.pass}</span></div>}
                  {data.fail > 0 && <div style={{ width:`${(data.fail/data.total)*100}%`, background:T.red, display:"flex", alignItems:"center", justifyContent:"center" }}><span style={{ fontSize:9, color:"white", fontWeight:700 }}>{data.fail}</span></div>}
                </div>
              </div>
            );
          })}
          <div style={{ marginTop:16, textAlign:"center" }}>
            <DonutChart passed={passed} failed={failed} other={total - passed - failed} total={total} T={T} />
          </div>
        </div>
      </div>

      {/* Failed cases */}
      {failed > 0 && (
        <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:14, padding:18, marginBottom:14 }}>
          <h4 style={{ margin:"0 0 12px", fontSize:14, fontWeight:600, color:T.red }}>Failed Test Cases ({failed})</h4>
          {cases.filter(c => c.status === "fail").map(c => {
            const tc  = state.testcases.find(t => t.id === c.tcId); if (!tc) return null;
            const mod = state.modules.find(m => m.id === tc.moduleId);
            return (
              <div key={c.tcId} style={{ borderLeft:`3px solid ${T.red}`, paddingLeft:12, marginBottom:12, paddingBottom:12, borderBottom:`1px solid ${T.border}` }}>
                <div style={{ display:"flex", gap:8, alignItems:"center", marginBottom:4, flexWrap:"wrap" }}>
                  <span style={{ fontSize:11, fontFamily:"monospace", color:T.blue, background:T.blueSoft, padding:"1px 7px", borderRadius:4 }}>{tc.tcId}</span>
                  <PriBadge p={tc.priority} T={T} /><span style={{ fontSize:11, color:T.textMuted }}>{mod?.name}</span>
                </div>
                <div style={{ fontSize:13, fontWeight:500, color:T.text }}>{tc.title}</div>
                {c.comment && <div style={{ fontSize:12, color:T.textMuted, marginTop:4, fontStyle:"italic" }}>"{c.comment}"</div>}
              </div>
            );
          })}
        </div>
      )}

      {/* Full table */}
      <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:14, padding:18 }}>
        <h4 style={{ margin:"0 0 12px", fontSize:14, fontWeight:600, color:T.textSec }}>All Test Cases</h4>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead>
              <tr style={{ background:T.surface2 }}>
                {["TC ID","Title","Module","Priority","Status","Comment"].map(hd => <th key={hd} style={{ padding:"8px 10px", textAlign:"left", color:T.textMuted, fontWeight:600, borderBottom:`1px solid ${T.border}`, fontSize:11, textTransform:"uppercase", letterSpacing:0.4, whiteSpace:"nowrap" }}>{hd}</th>)}
              </tr>
            </thead>
            <tbody>
              {cases.map(c => {
                const tc  = state.testcases.find(t => t.id === c.tcId); if (!tc) return null;
                const mod = state.modules.find(m => m.id === tc.moduleId);
                return (
                  <tr key={c.tcId} style={{ borderBottom:`1px solid ${T.border}` }}>
                    <td style={{ padding:"7px 10px", fontFamily:"monospace", color:T.blue, fontSize:11, whiteSpace:"nowrap" }}>{tc.tcId}</td>
                    <td style={{ padding:"7px 10px", color:T.text, maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{tc.title}</td>
                    <td style={{ padding:"7px 10px", color:T.textMuted, whiteSpace:"nowrap" }}>{mod?.name || "—"}</td>
                    <td style={{ padding:"7px 10px" }}><PriBadge p={tc.priority} T={T} /></td>
                    <td style={{ padding:"7px 10px" }}><StatusBadge status={c.status} T={T} small /></td>
                    <td style={{ padding:"7px 10px", color:T.textMuted, fontStyle:"italic", maxWidth:140, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{c.comment || "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ReportsView({ state, T }) {
  const completedRuns = state.runs.filter(r => r.status === "completed");
  const [selRun, setSelRun] = useState(() => completedRuns[0]?.id || null);
  const run = state.runs.find(r => r.id === selRun);

  const exportExcel = (run, state) => {
    const wb  = XLSX.utils.book_new();
    const rows = run.cases.map(c => {
      const tc   = state.testcases.find(t => t.id === c.tcId) || {};
      const mod  = state.modules.find(m => m.id === tc.moduleId) || {};
      const proj = state.projects.find(p => p.id === tc.projectId) || {};
      return { "Run Name":run.name, "TC ID":tc.tcId, "Title":tc.title, "Module":mod.name, "Project":proj.name, "Priority":tc.priority, "Status":c.status, "Comment":c.comment, "Executed At":c.executedAt };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Results");
    XLSX.writeFile(wb, `Report_${run.name.replace(/\s+/g,"_")}.xlsx`);
  };

  if (completedRuns.length === 0) return (
    <div style={{ padding:40, textAlign:"center", color:T.textMuted }}>
      <div style={{ fontSize:40, marginBottom:12 }}>📊</div>
      <div style={{ fontSize:15, fontWeight:500, color:T.textSec }}>No completed runs yet</div>
      <div style={{ fontSize:13, marginTop:4 }}>Complete a test run to see reports</div>
    </div>
  );

  return (
    <div style={{ padding:20, maxWidth:1000, margin:"0 auto", overflowY:"auto", height:"100%" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:18, flexWrap:"wrap", gap:10 }}>
        <h2 style={{ margin:0, fontSize:18, fontWeight:700, color:T.text }}>Reports</h2>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <TfSelect T={T} value={selRun || ""} onChange={e => setSelRun(e.target.value)} style={{ width:220 }}>
            {completedRuns.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
          </TfSelect>
          {run && <Btn variant="outline" T={T} onClick={() => exportExcel(run, state)} style={{ fontSize:12 }}>↓ Excel</Btn>}
        </div>
      </div>
      {run && <ReportContent run={run} state={state} T={T} />}
    </div>
  );
}

// ─── Excel Import / Export ────────────────────────────────────────────────────
function exportToExcel(state) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.projects.map(p => ({ ID:p.id, Name:p.name, Key:p.key, Description:p.description, CreatedAt:p.createdAt }))), "Projects");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.modules.map(m => ({ ID:m.id, ProjectID:m.projectId, Name:m.name, Key:m.key, Description:m.description }))), "Modules");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.testcases.map(tc => ({ ID:tc.id, ProjectID:tc.projectId, ModuleID:tc.moduleId, TCID:tc.tcId, Num:tc.num, Title:tc.title, Description:tc.description, Steps:tc.steps, ExpectedResult:tc.expectedResult, Priority:tc.priority, Tags:tc.tags, Status:tc.status }))), "TestCases");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.suites.map(s => ({ ID:s.id, Name:s.name, Description:s.description, CreatedAt:s.createdAt, Inclusions:JSON.stringify(s.inclusions) }))), "Suites");
  const runRows = [];
  state.runs.forEach(r => r.cases.forEach(c => runRows.push({ RunID:r.id, RunName:r.name, RunStatus:r.status, CreatedAt:r.createdAt, CompletedAt:r.completedAt||"", TCID:c.tcId, CaseStatus:c.status, Comment:c.comment, ExecutedAt:c.executedAt||"" })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(runRows.length ? runRows : [{ note:"No runs yet" }]), "Runs");
  XLSX.writeFile(wb, "TestFlow_Export.xlsx");
}

function importFromExcel(file, dispatch, onToast) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type:"binary" });
      const gs = name => { const ws = wb.Sheets[name]; return ws ? XLSX.utils.sheet_to_json(ws) : []; };
      const projects  = gs("Projects").map(r  => ({ id:r.ID||uid(), name:r.Name, key:r.Key, description:r.Description||"", createdAt:r.CreatedAt||new Date().toISOString() }));
      const modules   = gs("Modules").map(r   => ({ id:r.ID||uid(), projectId:r.ProjectID, name:r.Name, key:r.Key, description:r.Description||"" }));
      const testcases = gs("TestCases").map(r => ({ id:r.ID||uid(), projectId:r.ProjectID, moduleId:r.ModuleID, tcId:r.TCID, num:r.Num||1, title:r.Title, description:r.Description||"", steps:r.Steps||"", expectedResult:r.ExpectedResult||"", priority:r.Priority||"medium", tags:r.Tags||"", status:r.Status||"active" }));
      const suites    = gs("Suites").map(r    => ({ id:r.ID||uid(), name:r.Name, description:r.Description||"", createdAt:r.CreatedAt||new Date().toISOString(), inclusions:(()=>{ try { return JSON.parse(r.Inclusions||"[]"); } catch { return []; } })() }));
      const runMap    = {};
      gs("Runs").forEach(r => {
        if (!r.RunID) return;
        if (!runMap[r.RunID]) runMap[r.RunID] = { id:r.RunID, name:r.RunName, description:"", suiteIds:[], cases:[], status:r.RunStatus||"active", createdAt:r.CreatedAt||new Date().toISOString(), completedAt:r.CompletedAt||null };
        runMap[r.RunID].cases.push({ tcId:r.TCID, status:r.CaseStatus||"pending", comment:r.Comment||"", executedAt:r.ExecutedAt||null });
      });
      dispatch({ type:"IMPORT", data:{ projects, modules, testcases, suites, runs:Object.values(runMap) } });
      onToast("Excel imported successfully!", "success");
    } catch (err) {
      onToast("Import failed: " + err.message, "error");
    }
  };
  reader.readAsBinaryString(file);
}

// ─── Google Sheets Panel (shown in top bar) ───────────────────────────────────
function GSheetPanel({ gs, state, dispatch, T, isMobile }) {
  const { gsUser, gsStatus, signIn, signOut, saveToSheets, loadFromSheets, openSheet } = gs;

  const handleLoad = async () => {
    const data = await loadFromSheets();
    if (data) dispatch({ type:"IMPORT", data });
  };

  if (gsStatus === "idle" || gsStatus === "signing-in") {
    return (
      <Btn variant="outline" T={T} onClick={signIn} disabled={gsStatus === "signing-in"} style={{ fontSize:12, gap:5 }}>
        <span style={{ fontSize:14 }}>🔗</span>
        {isMobile ? "" : gsStatus === "signing-in" ? "Signing in…" : "Connect Google Sheets"}
      </Btn>
    );
  }

  return (
    <div style={{ display:"flex", alignItems:"center", gap:5 }}>
      {!isMobile && (
        <div style={{ display:"flex", alignItems:"center", gap:6, background:T.greenSoft, borderRadius:8, padding:"4px 10px" }}>
          {gsUser?.picture && <img src={gsUser.picture} alt="" style={{ width:18, height:18, borderRadius:"50%" }} referrerPolicy="no-referrer" />}
          <span style={{ fontSize:11, color:T.greenText, fontWeight:600, maxWidth:120, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{gsUser?.email}</span>
        </div>
      )}
      <Btn variant="outline" T={T} onClick={() => saveToSheets(state)} disabled={gsStatus === "saving"} style={{ fontSize:12 }} title="Save to Google Sheets">
        {gsStatus === "saving" ? "Saving…" : isMobile ? "☁↑" : "☁ Save"}
      </Btn>
      <Btn variant="outline" T={T} onClick={handleLoad} disabled={gsStatus === "loading"} style={{ fontSize:12 }} title="Load from Google Sheets">
        {gsStatus === "loading" ? "Loading…" : isMobile ? "☁↓" : "☁ Load"}
      </Btn>
      {!isMobile && <Btn variant="ghost" T={T} onClick={openSheet} style={{ fontSize:12 }} title="Open Sheet in browser">📄</Btn>}
      <Btn variant="ghost" T={T} onClick={signOut} style={{ fontSize:11, color:T.textMuted }} title="Sign out">⏏</Btn>
    </div>
  );
}

// ─── Navigation ───────────────────────────────────────────────────────────────
const NAV_ITEMS = [
  { id:"projects", label:"Projects",    icon:"📁" },
  { id:"suites",   label:"Suites",      icon:"📋" },
  { id:"runs",     label:"Runs",        icon:"▶"  },
  { id:"reports",  label:"Reports",     icon:"📊" },
];

// ─── Root App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [state, dispatchRaw] = useReducer(reducer, null, makeInitial);
  const [nav, setNav]         = useState("projects");
  const [dark, setDark]       = useState(false);
  const [isMobile, setIsMobile] = useState(window.innerWidth <= 700);
  const [toast, setToast]     = useState(null);
  const toastTimer            = useRef(null);
  const fileRef               = useRef();
  const T = useMemo(() => makeTheme(dark), [dark]);

  // Responsive detection
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth <= 700);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  // Toast helper
  const showToast = useCallback((message, type = "info") => {
    setToast({ message, type });
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 3500);
  }, []);

  const dispatch = useCallback(action => dispatchRaw(action), []);

  // Google Sheets hook
  const gs = useGoogleSheets(showToast);

  const totalTCs      = state.testcases.length;
  const completedRuns = state.runs.filter(r => r.status === "completed").length;

  // Global CSS (injected once)
  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `
      *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
      body { overflow: hidden; }
      ::-webkit-scrollbar { width: 5px; height: 5px; }
      ::-webkit-scrollbar-track { background: transparent; }
      ::-webkit-scrollbar-thumb { border-radius: 3px; background: #D1D5DB; }
      .fade-in { animation: fadeIn 0.18s ease; }
      @keyframes fadeIn { from { opacity:0; transform:translateY(3px); } to { opacity:1; transform:translateY(0); } }
      .panel-slide { transition: width 0.25s cubic-bezier(.4,0,.2,1), min-width 0.25s cubic-bezier(.4,0,.2,1); }
    `;
    document.head.appendChild(style);
    return () => document.head.removeChild(style);
  }, []);

  return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif", height:"100vh", display:"flex", flexDirection:"column", background:T.bg, color:T.text }}>

      {/* ── Top Bar ── */}
      <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:isMobile ? "0 10px" : "0 18px", height:isMobile ? 48 : 54, display:"flex", alignItems:"center", gap:0, flexShrink:0, zIndex:10 }}>
        {/* Logo */}
        <div style={{ display:"flex", alignItems:"center", gap:8, marginRight:isMobile ? 10 : 20 }}>
          <div style={{ width:28, height:28, background:T.blue, borderRadius:7, display:"flex", alignItems:"center", justifyContent:"center" }}>
            <span style={{ fontSize:13, color:"white", fontWeight:700 }}>✓</span>
          </div>
          {!isMobile && <span style={{ fontSize:15, fontWeight:700, color:T.text }}>TestFlow</span>}
        </div>

        {/* Nav */}
        <nav style={{ display:"flex", gap:2, flex:1 }}>
          {NAV_ITEMS.map(item => (
            <button key={item.id} onClick={() => setNav(item.id)}
              style={{ background:nav === item.id ? T.navActive : "transparent", border:"none", borderRadius:8, padding:isMobile ? "5px 9px" : "6px 13px", cursor:"pointer", fontFamily:"inherit", fontSize:isMobile ? 11 : 13, fontWeight:nav === item.id ? 600 : 500, color:nav === item.id ? T.blue : T.textMuted, display:"flex", alignItems:"center", gap:4, outline:"none" }}>
              <span>{item.icon}</span>{!isMobile && <span> {item.label}</span>}
            </button>
          ))}
        </nav>

        {/* Stats */}
        {!isMobile && (
          <div style={{ display:"flex", gap:12, fontSize:12, color:T.textMuted, marginRight:14 }}>
            <span>{state.projects.length} projects</span>
            <span>{totalTCs} tests</span>
            <span>{completedRuns}/{state.runs.length} runs done</span>
          </div>
        )}

        {/* Actions */}
        <div style={{ display:"flex", gap:5, alignItems:"center" }}>
          {/* Google Sheets */}
          <GSheetPanel gs={gs} state={state} dispatch={dispatch} T={T} isMobile={isMobile} />

          {/* Divider */}
          <div style={{ width:1, height:20, background:T.border, margin:"0 4px" }} />

          {/* Excel export/import */}
          <Btn variant="outline" T={T} style={{ fontSize:12, padding:"5px 10px" }} onClick={() => exportToExcel(state)} title="Export to Excel">
            {isMobile ? "↑" : "↑ Excel"}
          </Btn>
          <Btn variant="outline" T={T} style={{ fontSize:12, padding:"5px 10px" }} onClick={() => fileRef.current.click()} title="Import from Excel">
            {isMobile ? "↓" : "↓ Excel"}
          </Btn>
          <input ref={fileRef} type="file" accept=".xlsx" style={{ display:"none" }} onChange={e => { if (e.target.files[0]) importFromExcel(e.target.files[0], dispatch, showToast); e.target.value = ""; }} />

          {/* Divider */}
          <div style={{ width:1, height:20, background:T.border, margin:"0 4px" }} />

          {/* Dark mode toggle */}
          <Btn variant="ghost" T={T} onClick={() => setDark(d => !d)} style={{ padding:"5px 8px", fontSize:16 }} title={dark ? "Light mode" : "Dark mode"}>
            {dark ? "☀" : "🌙"}
          </Btn>
        </div>
      </div>

      {/* ── Main Content ── */}
      <div style={{ flex:1, overflow:"hidden" }}>
        {nav === "projects" && <ProjectsView state={state} dispatch={dispatch} T={T} isMobile={isMobile} />}
        {nav === "suites"   && <div style={{ height:"100%", overflowY:"auto" }}><SuitesView state={state} dispatch={dispatch} T={T} /></div>}
        {nav === "runs"     && <div style={{ height:"100%", overflowY:"auto" }}><RunsView state={state} dispatch={dispatch} T={T} isMobile={isMobile} /></div>}
        {nav === "reports"  && <div style={{ height:"100%", overflowY:"auto" }}><ReportsView state={state} T={T} /></div>}
      </div>

      {/* ── Footer ── */}
      <div style={{ background:T.surface, borderTop:`1px solid ${T.border}`, padding:"5px 18px", display:"flex", justifyContent:"center", alignItems:"center", flexShrink:0 }}>
        <span style={{ fontSize:11, color:T.textMuted }}>
          Built by <strong style={{ color:T.textSec }}>Joseph</strong>, <strong style={{ color:T.textSec }}>Akhil</strong>
        </span>
      </div>

      {/* ── Toast ── */}
      {toast && <Toast message={toast.message} type={toast.type} T={T} />}
    </div>
  );
}