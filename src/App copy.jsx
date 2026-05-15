import { useState, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import logo1 from "./logo1.png";

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  blue: "#1A6BCC",
  blueSoft: "#E8F1FB",
  blueHover: "#1559AB",
  green: "#1A7F4B",
  greenSoft: "#E6F4EC",
  red: "#C53030",
  redSoft: "#FDECEA",
  amber: "#B45309",
  amberSoft: "#FEF3C7",
  gray50: "#F9FAFB",
  gray100: "#F3F4F6",
  gray200: "#E5E7EB",
  gray300: "#D1D5DB",
  gray400: "#9CA3AF",
  gray500: "#6B7280",
  gray600: "#4B5563",
  gray700: "#374151",
  gray800: "#1F2937",
  white: "#FFFFFF",
  border: "#E5E7EB",
  purple: "#6D28D9",
  purpleSoft: "#EDE9FE",
  teal: "#0F766E",
  tealSoft: "#CCFBF1",
};

const STATUS_META = {
  pass: { label: "Pass", color: C.green, bg: C.greenSoft },
  fail: { label: "Fail", color: C.red, bg: C.redSoft },
  pending: { label: "Pending", color: C.gray500, bg: C.gray100 },
  blocked: { label: "Blocked", color: C.amber, bg: C.amberSoft },
  skipped: { label: "Skipped", color: C.purple, bg: C.purpleSoft },
  "in-progress": { label: "In Progress", color: C.blue, bg: C.blueSoft },
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function genTCId(projectKey, moduleKey, num) {
  return `${projectKey}-${moduleKey}-TC${String(num).padStart(3, "0")}`;
}

function StatusBadge({ status, small }) {
  const m = STATUS_META[status] || STATUS_META.pending;
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        padding: small ? "2px 8px" : "3px 10px",
        borderRadius: 20,
        fontSize: small ? 11 : 12,
        fontWeight: 600,
        color: m.color,
        background: m.bg,
        whiteSpace: "nowrap",
      }}
    >
      {m.label}
    </span>
  );
}

function Btn({ onClick, variant = "ghost", children, style, disabled, title }) {
  const base = {
    border: "1px solid",
    borderRadius: 8,
    cursor: disabled ? "not-allowed" : "pointer",
    fontFamily: "inherit",
    fontSize: 13,
    fontWeight: 500,
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    padding: "6px 14px",
    transition: "all 0.15s",
    opacity: disabled ? 0.5 : 1,
  };
  const variants = {
    primary: { background: C.blue, color: C.white, borderColor: C.blue },
    danger: { background: C.red, color: C.white, borderColor: C.red },
    success: { background: C.green, color: C.white, borderColor: C.green },
    outline: { background: C.white, color: C.gray700, borderColor: C.border },
    ghost: { background: "transparent", color: C.gray600, borderColor: "transparent" },
  };
  return (
    <button
      onClick={disabled ? undefined : onClick}
      title={title}
      style={{ ...base, ...variants[variant], ...style }}
    >
      {children}
    </button>
  );
}

function Modal({ title, onClose, children, width = 560 }) {
  return (
    <div
      style={{
        position: "fixed", inset: 0, zIndex: 1000,
        background: "rgba(15,23,42,0.45)",
        display: "flex", alignItems: "center", justifyContent: "center",
        padding: 16,
      }}
      onClick={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        style={{
          background: C.white, borderRadius: 14,
          width: "100%", maxWidth: width,
          maxHeight: "90vh", overflow: "auto",
          boxShadow: "0 20px 60px rgba(0,0,0,0.15)",
        }}
      >
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "18px 24px", borderBottom: `1px solid ${C.border}`,
          position: "sticky", top: 0, background: C.white, zIndex: 1,
        }}>
          <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600, color: C.gray800 }}>{title}</h2>
          <Btn onClick={onClose} variant="ghost" style={{ padding: "4px 8px" }}>✕</Btn>
        </div>
        <div style={{ padding: "20px 24px" }}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children, required }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <label style={{ display: "block", fontSize: 13, fontWeight: 500, color: C.gray700, marginBottom: 6 }}>
        {label}{required && <span style={{ color: C.red }}> *</span>}
      </label>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%", boxSizing: "border-box",
  border: `1px solid ${C.border}`, borderRadius: 8,
  padding: "8px 12px", fontSize: 13, fontFamily: "inherit",
  color: C.gray800, background: C.white, outline: "none",
};

function Input({ ...props }) {
  return <input style={inputStyle} {...props} />;
}
function Textarea({ ...props }) {
  return <textarea style={{ ...inputStyle, resize: "vertical", minHeight: 80 }} {...props} />;
}
function Select({ children, ...props }) {
  return <select style={{ ...inputStyle, cursor: "pointer" }} {...props}>{children}</select>;
}

// ─── Initial State ────────────────────────────────────────────────────────────
function makeInitialState() {
  const proj1Id = uid(), proj2Id = uid();
  const mod1Id = uid(), mod2Id = uid(), mod3Id = uid();
  const tc1 = uid(), tc2 = uid(), tc3 = uid(), tc4 = uid();

  const projects = [
    { id: proj1Id, name: "E-Commerce Platform", key: "ECP", description: "Online shopping platform", createdAt: new Date().toISOString() },
    { id: proj2Id, name: "Admin Dashboard", key: "ADM", description: "Internal admin tool", createdAt: new Date().toISOString() },
  ];

  const modules = [
    { id: mod1Id, projectId: proj1Id, name: "Authentication", key: "AUTH", description: "Login, register, session" },
    { id: mod2Id, projectId: proj1Id, name: "Cart & Checkout", key: "CART", description: "Shopping cart and payment" },
    { id: mod3Id, projectId: proj2Id, name: "User Management", key: "USR", description: "CRUD for users" },
  ];

  const testcases = [
    {
      id: tc1, moduleId: mod1Id, projectId: proj1Id,
      tcId: "ECP-AUTH-TC001", title: "Valid login with correct credentials",
      description: "Verify user can login with correct email and password",
      steps: "1. Navigate to /login\n2. Enter valid email\n3. Enter valid password\n4. Click Sign In",
      expectedResult: "User is redirected to dashboard. Session token is set.",
      priority: "high", status: "active", num: 1, tags: "smoke,regression",
    },
    {
      id: tc2, moduleId: mod1Id, projectId: proj1Id,
      tcId: "ECP-AUTH-TC002", title: "Login with invalid password shows error",
      description: "Verify appropriate error shown for wrong password",
      steps: "1. Navigate to /login\n2. Enter valid email\n3. Enter wrong password\n4. Click Sign In",
      expectedResult: "Error message: 'Invalid credentials'. User stays on login page.",
      priority: "medium", status: "active", num: 2, tags: "regression",
    },
    {
      id: tc3, moduleId: mod2Id, projectId: proj1Id,
      tcId: "ECP-CART-TC001", title: "Add item to cart",
      description: "Verify product can be added to shopping cart",
      steps: "1. Browse to product page\n2. Click 'Add to Cart'\n3. View cart",
      expectedResult: "Item appears in cart with correct price and quantity.",
      priority: "high", status: "active", num: 1, tags: "smoke",
    },
    {
      id: tc4, moduleId: mod3Id, projectId: proj2Id,
      tcId: "ADM-USR-TC001", title: "Create new user account",
      description: "Admin can create a new user from the management panel",
      steps: "1. Login as admin\n2. Go to Users\n3. Click Create User\n4. Fill form\n5. Submit",
      expectedResult: "New user appears in list. Welcome email sent.",
      priority: "high", status: "active", num: 1, tags: "smoke",
    },
  ];

  return { projects, modules, testcases, suites: [], runs: [] };
}

// ─── Nav ──────────────────────────────────────────────────────────────────────
const NAV_ITEMS = [
  { id: "projects", label: "Projects", icon: "📁" },
  { id: "suites", label: "Test Suites", icon: "📋" },
  { id: "runs", label: "Test Runs", icon: "▶" },
  { id: "reports", label: "Reports", icon: "📊" },
];

// ─── Projects View ────────────────────────────────────────────────────────────
function ProjectsView({ state, dispatch }) {
  const { projects, modules, testcases } = state;
  const [selected, setSelected] = useState(null);
  const [selectedModule, setSelectedModule] = useState(null);
  const [showAddProject, setShowAddProject] = useState(false);
  const [showAddModule, setShowAddModule] = useState(false);
  const [showAddTC, setShowAddTC] = useState(false);
  const [editTC, setEditTC] = useState(null);
  const [viewTC, setViewTC] = useState(null);

  const project = projects.find((p) => p.id === selected);
  const projectModules = modules.filter((m) => m.projectId === selected);
  const module = projectModules.find((m) => m.id === selectedModule);
  const moduleTCs = testcases.filter((tc) => tc.moduleId === selectedModule);

  return (
    <div style={{ display: "flex", height: "100%", overflow: "hidden" }}>
      {/* Projects panel */}
      <div style={{
        width: 220, flexShrink: 0, borderRight: `1px solid ${C.border}`,
        background: C.gray50, overflowY: "auto", display: "flex", flexDirection: "column",
      }}>
        <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 12, fontWeight: 600, color: C.gray500, textTransform: "uppercase", letterSpacing: 0.5 }}>Projects</span>
          <Btn onClick={() => setShowAddProject(true)} variant="primary" style={{ padding: "3px 8px", fontSize: 11 }}>+ New</Btn>
        </div>
        {projects.map((p) => (
          <div
            key={p.id}
            onClick={() => { setSelected(p.id); setSelectedModule(null); }}
            style={{
              padding: "10px 16px", cursor: "pointer", borderBottom: `1px solid ${C.border}`,
              background: selected === p.id ? C.blueSoft : "transparent",
              borderLeft: selected === p.id ? `3px solid ${C.blue}` : "3px solid transparent",
            }}
          >
            <div style={{ fontSize: 13, fontWeight: 600, color: selected === p.id ? C.blue : C.gray800 }}>{p.name}</div>
            <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>{p.key}</div>
          </div>
        ))}
      </div>

      {/* Modules panel */}
      <div style={{
        width: 200, flexShrink: 0, borderRight: `1px solid ${C.border}`,
        background: C.white, overflowY: "auto", display: "flex", flexDirection: "column",
      }}>
        {project ? (
          <>
            <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 12, fontWeight: 600, color: C.gray500, textTransform: "uppercase", letterSpacing: 0.5 }}>Modules</span>
              <Btn onClick={() => setShowAddModule(true)} variant="primary" style={{ padding: "3px 8px", fontSize: 11 }}>+ New</Btn>
            </div>
            {projectModules.map((m) => (
              <div
                key={m.id}
                onClick={() => setSelectedModule(m.id)}
                style={{
                  padding: "10px 16px", cursor: "pointer", borderBottom: `1px solid ${C.border}`,
                  background: selectedModule === m.id ? C.blueSoft : "transparent",
                  borderLeft: selectedModule === m.id ? `3px solid ${C.blue}` : "3px solid transparent",
                }}
              >
                <div style={{ fontSize: 13, fontWeight: 500, color: selectedModule === m.id ? C.blue : C.gray700 }}>{m.name}</div>
                <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>
                  {testcases.filter((tc) => tc.moduleId === m.id).length} tests
                </div>
              </div>
            ))}
            {projectModules.length === 0 && (
              <div style={{ padding: 16, color: C.gray400, fontSize: 13 }}>No modules yet</div>
            )}
          </>
        ) : (
          <div style={{ padding: 24, color: C.gray400, fontSize: 13 }}>Select a project</div>
        )}
      </div>

      {/* Test cases panel */}
      <div style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column" }}>
        {module ? (
          <>
            <div style={{
              padding: "14px 20px", borderBottom: `1px solid ${C.border}`,
              display: "flex", justifyContent: "space-between", alignItems: "center",
              background: C.white, position: "sticky", top: 0, zIndex: 1,
            }}>
              <div>
                <div style={{ fontSize: 15, fontWeight: 600, color: C.gray800 }}>{module.name}</div>
                <div style={{ fontSize: 12, color: C.gray400 }}>{module.description}</div>
              </div>
              <Btn onClick={() => setShowAddTC(true)} variant="primary">+ New Test Case</Btn>
            </div>
            <div style={{ padding: 16 }}>
              {moduleTCs.length === 0 && (
                <div style={{ textAlign: "center", padding: 40, color: C.gray400 }}>No test cases yet. Add one!</div>
              )}
              {moduleTCs.map((tc) => (
                <div key={tc.id} style={{
                  background: C.white, border: `1px solid ${C.border}`, borderRadius: 10,
                  padding: "14px 16px", marginBottom: 8,
                  display: "flex", alignItems: "flex-start", gap: 12,
                }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 11, fontWeight: 700, color: C.blue, background: C.blueSoft, padding: "2px 8px", borderRadius: 6, fontFamily: "monospace" }}>{tc.tcId}</span>
                      <PriorityBadge priority={tc.priority} />
                    </div>
                    <div style={{ fontSize: 14, fontWeight: 500, color: C.gray800, marginTop: 6 }}>{tc.title}</div>
                    <div style={{ fontSize: 12, color: C.gray500, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tc.description}</div>
                  </div>
                  <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                    <Btn onClick={() => setViewTC(tc)} variant="outline" style={{ padding: "4px 10px", fontSize: 12 }}>View</Btn>
                    <Btn onClick={() => setEditTC(tc)} variant="outline" style={{ padding: "4px 10px", fontSize: 12 }}>Edit</Btn>
                    <Btn onClick={() => dispatch({ type: "DELETE_TC", id: tc.id })} variant="ghost" style={{ padding: "4px 8px", fontSize: 12, color: C.red }}>✕</Btn>
                  </div>
                </div>
              ))}
            </div>
          </>
        ) : project ? (
          <div style={{ padding: 40, color: C.gray400, fontSize: 14, textAlign: "center" }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>📂</div>
            Select a module to view test cases
          </div>
        ) : (
          <div style={{ padding: 40, color: C.gray400, fontSize: 14, textAlign: "center" }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>📁</div>
            Select a project to get started
          </div>
        )}
      </div>

      {/* Modals */}
      {showAddProject && (
        <AddProjectModal onClose={() => setShowAddProject(false)} onSave={(p) => { dispatch({ type: "ADD_PROJECT", project: p }); setShowAddProject(false); }} />
      )}
      {showAddModule && project && (
        <AddModuleModal project={project} onClose={() => setShowAddModule(false)} onSave={(m) => { dispatch({ type: "ADD_MODULE", module: m }); setShowAddModule(false); }} />
      )}
      {(showAddTC || editTC) && module && (
        <AddTCModal
          module={module} project={project}
          existing={editTC}
          tcCount={moduleTCs.length}
          onClose={() => { setShowAddTC(false); setEditTC(null); }}
          onSave={(tc) => {
            dispatch({ type: editTC ? "EDIT_TC" : "ADD_TC", tc });
            setShowAddTC(false); setEditTC(null);
          }}
        />
      )}
      {viewTC && (
        <ViewTCModal tc={viewTC} onClose={() => setViewTC(null)} onEdit={() => { setEditTC(viewTC); setViewTC(null); }} />
      )}
    </div>
  );
}

function PriorityBadge({ priority }) {
  const meta = { high: { color: C.red, bg: C.redSoft }, medium: { color: C.amber, bg: C.amberSoft }, low: { color: C.green, bg: C.greenSoft } };
  const m = meta[priority] || meta.medium;
  return (
    <span style={{ fontSize: 10, fontWeight: 700, color: m.color, background: m.bg, padding: "1px 7px", borderRadius: 20, textTransform: "uppercase", letterSpacing: 0.5 }}>{priority}</span>
  );
}

function AddProjectModal({ onClose, onSave }) {
  const [form, setForm] = useState({ name: "", key: "", description: "" });
  const f = (k) => (e) => setForm((p) => ({ ...p, [k]: e.target.value }));
  return (
    <Modal title="New Project" onClose={onClose} width={480}>
      <Field label="Project Name" required><Input value={form.name} onChange={f("name")} placeholder="E.g. Mobile App" /></Field>
      <Field label="Key (short prefix)" required><Input value={form.key} onChange={f("key")} placeholder="E.g. APP" maxLength={5} style={{ textTransform: "uppercase" }} /></Field>
      <Field label="Description"><Textarea value={form.description} onChange={f("description")} /></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn onClick={onClose} variant="outline">Cancel</Btn>
        <Btn variant="primary" onClick={() => { if (!form.name || !form.key) return; onSave({ id: uid(), ...form, key: form.key.toUpperCase(), createdAt: new Date().toISOString() }); }}>Create Project</Btn>
      </div>
    </Modal>
  );
}

function AddModuleModal({ project, onClose, onSave }) {
  const [form, setForm] = useState({ name: "", key: "", description: "" });
  const f = (k) => (e) => setForm((p) => ({ ...p, [k]: e.target.value }));
  return (
    <Modal title={`New Module — ${project.name}`} onClose={onClose} width={480}>
      <Field label="Module Name" required><Input value={form.name} onChange={f("name")} placeholder="E.g. Authentication" /></Field>
      <Field label="Key" required><Input value={form.key} onChange={f("key")} placeholder="E.g. AUTH" maxLength={8} style={{ textTransform: "uppercase" }} /></Field>
      <Field label="Description"><Textarea value={form.description} onChange={f("description")} /></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn onClick={onClose} variant="outline">Cancel</Btn>
        <Btn variant="primary" onClick={() => { if (!form.name || !form.key) return; onSave({ id: uid(), projectId: project.id, ...form, key: form.key.toUpperCase() }); }}>Create Module</Btn>
      </div>
    </Modal>
  );
}

function AddTCModal({ module, project, existing, tcCount, onClose, onSave }) {
  const nextNum = existing ? existing.num : tcCount + 1;
  const [form, setForm] = useState(existing || {
    title: "", description: "", steps: "", expectedResult: "",
    priority: "medium", tags: "",
  });
  const f = (k) => (e) => setForm((p) => ({ ...p, [k]: e.target.value }));
  const tcId = existing?.tcId || genTCId(project.key, module.key, nextNum);

  return (
    <Modal title={existing ? "Edit Test Case" : "New Test Case"} onClose={onClose} width={640}>
      <div style={{ background: C.blueSoft, borderRadius: 8, padding: "8px 12px", marginBottom: 16, fontSize: 12, color: C.blue, fontWeight: 600, fontFamily: "monospace" }}>
        ID: {tcId}
      </div>
      <Field label="Title" required><Input value={form.title} onChange={f("title")} placeholder="Short descriptive title" /></Field>
      <Field label="Description"><Textarea value={form.description} onChange={f("description")} placeholder="What does this test verify?" /></Field>
      <Field label="Test Steps" required>
        <Textarea value={form.steps} onChange={f("steps")} placeholder="1. Step one&#10;2. Step two&#10;3. Step three" style={{ minHeight: 100 }} />
      </Field>
      <Field label="Expected Result" required><Textarea value={form.expectedResult} onChange={f("expectedResult")} placeholder="What should happen?" /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="Priority">
          <Select value={form.priority} onChange={f("priority")}>
            <option value="high">High</option>
            <option value="medium">Medium</option>
            <option value="low">Low</option>
          </Select>
        </Field>
        <Field label="Tags (comma separated)"><Input value={form.tags} onChange={f("tags")} placeholder="smoke, regression" /></Field>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn onClick={onClose} variant="outline">Cancel</Btn>
        <Btn variant="primary" onClick={() => {
          if (!form.title) return;
          onSave({
            ...(existing || {}),
            id: existing?.id || uid(),
            moduleId: module.id,
            projectId: project.id,
            tcId, num: nextNum,
            ...form,
            status: "active",
          });
        }}>{existing ? "Save Changes" : "Create Test Case"}</Btn>
      </div>
    </Modal>
  );
}

function ViewTCModal({ tc, onClose, onEdit }) {
  return (
    <Modal title={tc.tcId} onClose={onClose} width={600}>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <PriorityBadge priority={tc.priority} />
        {tc.tags && tc.tags.split(",").map((t) => (
          <span key={t} style={{ fontSize: 11, color: C.teal, background: C.tealSoft, padding: "2px 8px", borderRadius: 20 }}>{t.trim()}</span>
        ))}
      </div>
      <h3 style={{ margin: "0 0 8px", fontSize: 16, color: C.gray800 }}>{tc.title}</h3>
      {tc.description && <p style={{ color: C.gray500, fontSize: 13, margin: "0 0 16px" }}>{tc.description}</p>}
      <div style={{ background: C.gray50, borderRadius: 8, padding: 14, marginBottom: 12 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: C.gray400, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Test Steps</div>
        <pre style={{ margin: 0, fontSize: 13, color: C.gray700, whiteSpace: "pre-wrap", fontFamily: "inherit" }}>{tc.steps}</pre>
      </div>
      <div style={{ background: C.greenSoft, borderRadius: 8, padding: 14 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: C.green, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Expected Result</div>
        <div style={{ fontSize: 13, color: C.gray700 }}>{tc.expectedResult}</div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <Btn onClick={onClose} variant="outline">Close</Btn>
        <Btn onClick={onEdit} variant="primary">Edit</Btn>
      </div>
    </Modal>
  );
}

// ─── Test Suites View ─────────────────────────────────────────────────────────
function SuitesView({ state, dispatch }) {
  const { suites, projects, modules, testcases } = state;
  const [showAdd, setShowAdd] = useState(false);
  const [editSuite, setEditSuite] = useState(null);

  const tcCount = (suite) => resolveSuiteTCs(suite, state).length;

  return (
    <div style={{ padding: 24, maxWidth: 900, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: C.gray800 }}>Test Suites</h2>
          <p style={{ margin: "4px 0 0", color: C.gray400, fontSize: 13 }}>Group test cases into reusable suites</p>
        </div>
        <Btn variant="primary" onClick={() => setShowAdd(true)}>+ New Suite</Btn>
      </div>

      {suites.length === 0 && (
        <div style={{ textAlign: "center", padding: "60px 40px", background: C.gray50, borderRadius: 14, border: `1px dashed ${C.border}` }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>📋</div>
          <div style={{ fontSize: 15, fontWeight: 500, color: C.gray600, marginBottom: 4 }}>No test suites yet</div>
          <div style={{ fontSize: 13, color: C.gray400 }}>Create a suite to group test cases for a test run</div>
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {suites.map((suite) => (
          <div key={suite.id} style={{
            background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "16px 20px",
            display: "flex", alignItems: "center", gap: 16,
          }}>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: C.gray800, marginBottom: 4 }}>{suite.name}</div>
              {suite.description && <div style={{ fontSize: 12, color: C.gray400, marginBottom: 6 }}>{suite.description}</div>}
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                <span style={{ fontSize: 11, color: C.blue, background: C.blueSoft, padding: "2px 8px", borderRadius: 20 }}>{tcCount(suite)} test cases</span>
                {suite.inclusions.map((inc, i) => (
                  <span key={i} style={{ fontSize: 11, color: C.gray500, background: C.gray100, padding: "2px 8px", borderRadius: 20 }}>
                    {inc.type === "project" ? "Project: " : inc.type === "module" ? "Module: " : "TC: "}{inc.label}
                  </span>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <Btn variant="outline" style={{ fontSize: 12, padding: "5px 12px" }} onClick={() => setEditSuite(suite)}>Edit</Btn>
              <Btn variant="ghost" style={{ fontSize: 12, color: C.red }} onClick={() => dispatch({ type: "DELETE_SUITE", id: suite.id })}>✕</Btn>
            </div>
          </div>
        ))}
      </div>

      {(showAdd || editSuite) && (
        <AddSuiteModal
          existing={editSuite}
          state={state}
          onClose={() => { setShowAdd(false); setEditSuite(null); }}
          onSave={(s) => {
            dispatch({ type: editSuite ? "EDIT_SUITE" : "ADD_SUITE", suite: s });
            setShowAdd(false); setEditSuite(null);
          }}
        />
      )}
    </div>
  );
}

function resolveSuiteTCs(suite, state) {
  const { testcases, modules, projects } = state;
  const tcSet = new Set();
  suite.inclusions.forEach((inc) => {
    if (inc.type === "project") {
      testcases.filter((tc) => tc.projectId === inc.id).forEach((tc) => tcSet.add(tc.id));
    } else if (inc.type === "module") {
      testcases.filter((tc) => tc.moduleId === inc.id).forEach((tc) => tcSet.add(tc.id));
    } else if (inc.type === "tc") {
      tcSet.add(inc.id);
    }
  });
  return testcases.filter((tc) => tcSet.has(tc.id));
}

function AddSuiteModal({ existing, state, onClose, onSave }) {
  const { projects, modules, testcases } = state;
  const [form, setForm] = useState({ name: existing?.name || "", description: existing?.description || "" });
  const [inclusions, setInclusions] = useState(existing?.inclusions || []);
  const [addType, setAddType] = useState("project");
  const [addId, setAddId] = useState("");

  const getLabel = (type, id) => {
    if (type === "project") return projects.find((p) => p.id === id)?.name || id;
    if (type === "module") return modules.find((m) => m.id === id)?.name || id;
    if (type === "tc") return testcases.find((tc) => tc.id === id)?.tcId || id;
    return id;
  };

  const options = {
    project: projects,
    module: modules,
    tc: testcases,
  };

  const addInclusion = () => {
    if (!addId) return;
    if (inclusions.find((i) => i.type === addType && i.id === addId)) return;
    setInclusions((p) => [...p, { type: addType, id: addId, label: getLabel(addType, addId) }]);
    setAddId("");
  };

  return (
    <Modal title={existing ? "Edit Suite" : "New Test Suite"} onClose={onClose} width={580}>
      <Field label="Suite Name" required><Input value={form.name} onChange={(e) => setForm((p) => ({ ...p, name: e.target.value }))} /></Field>
      <Field label="Description"><Textarea value={form.description} onChange={(e) => setForm((p) => ({ ...p, description: e.target.value }))} /></Field>

      <div style={{ fontSize: 13, fontWeight: 600, color: C.gray700, marginBottom: 10 }}>Add Test Cases</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <Select value={addType} onChange={(e) => { setAddType(e.target.value); setAddId(""); }} style={{ width: 130 }}>
          <option value="project">Project</option>
          <option value="module">Module</option>
          <option value="tc">Test Case</option>
        </Select>
        <Select value={addId} onChange={(e) => setAddId(e.target.value)} style={{ flex: 1 }}>
          <option value="">Select…</option>
          {options[addType].map((item) => (
            <option key={item.id} value={item.id}>{item.name || item.title || item.tcId}</option>
          ))}
        </Select>
        <Btn variant="primary" onClick={addInclusion} style={{ padding: "6px 14px" }}>Add</Btn>
      </div>

      {inclusions.length > 0 && (
        <div style={{ background: C.gray50, borderRadius: 8, padding: 10, marginBottom: 16 }}>
          {inclusions.map((inc, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 0" }}>
              <span style={{ fontSize: 12, color: C.gray700 }}>
                <span style={{ color: C.gray400 }}>{inc.type}: </span>{inc.label}
              </span>
              <Btn variant="ghost" style={{ padding: "2px 6px", fontSize: 11, color: C.red }} onClick={() => setInclusions((p) => p.filter((_, j) => j !== i))}>✕</Btn>
            </div>
          ))}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn onClick={onClose} variant="outline">Cancel</Btn>
        <Btn variant="primary" onClick={() => {
          if (!form.name) return;
          onSave({ id: existing?.id || uid(), ...form, inclusions, createdAt: existing?.createdAt || new Date().toISOString() });
        }}>{existing ? "Save Changes" : "Create Suite"}</Btn>
      </div>
    </Modal>
  );
}

// ─── Test Runs View ───────────────────────────────────────────────────────────
function RunsView({ state, dispatch }) {
  const { runs, suites } = state;
  const [showAdd, setShowAdd] = useState(false);
  const [activeRun, setActiveRun] = useState(null);

  const run = runs.find((r) => r.id === activeRun);

  if (run) {
    return <RunExecutionView run={run} state={state} dispatch={dispatch} onBack={() => setActiveRun(null)} />;
  }

  return (
    <div style={{ padding: 24, maxWidth: 900, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: C.gray800 }}>Test Runs</h2>
          <p style={{ margin: "4px 0 0", color: C.gray400, fontSize: 13 }}>Execute and track test results</p>
        </div>
        <Btn variant="primary" onClick={() => setShowAdd(true)} disabled={suites.length === 0}>+ New Run</Btn>
      </div>

      {suites.length === 0 && (
        <div style={{ background: C.amberSoft, borderRadius: 10, padding: "12px 16px", marginBottom: 20, fontSize: 13, color: C.amber }}>
          ⚠ Create at least one test suite before starting a run.
        </div>
      )}

      {runs.length === 0 && (
        <div style={{ textAlign: "center", padding: "60px 40px", background: C.gray50, borderRadius: 14, border: `1px dashed ${C.border}` }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>▶</div>
          <div style={{ fontSize: 15, fontWeight: 500, color: C.gray600 }}>No test runs yet</div>
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {runs.map((r) => {
          const total = r.cases.length;
          const passed = r.cases.filter((c) => c.status === "pass").length;
          const failed = r.cases.filter((c) => c.status === "fail").length;
          const pending = r.cases.filter((c) => c.status === "pending").length;
          const pct = total ? Math.round((passed / total) * 100) : 0;
          return (
            <div key={r.id} style={{
              background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "16px 20px",
              display: "flex", alignItems: "center", gap: 16,
            }}>
              <div style={{ flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                  <span style={{ fontSize: 15, fontWeight: 600, color: C.gray800 }}>{r.name}</span>
                  <RunStatusBadge status={r.status} />
                </div>
                <div style={{ fontSize: 12, color: C.gray400, marginBottom: 8 }}>
                  {r.suiteIds.length} suite{r.suiteIds.length > 1 ? "s" : ""} · {total} test cases
                </div>
                <div style={{ display: "flex", gap: 12, fontSize: 12 }}>
                  <span style={{ color: C.green }}>✓ {passed} passed</span>
                  <span style={{ color: C.red }}>✗ {failed} failed</span>
                  <span style={{ color: C.gray400 }}>⋯ {pending} pending</span>
                  <span style={{ fontWeight: 600, color: C.blue }}>{pct}%</span>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                {r.status !== "completed" && (
                  <Btn variant="primary" style={{ fontSize: 12 }} onClick={() => setActiveRun(r.id)}>Execute</Btn>
                )}
                {r.status === "completed" && (
                  <Btn variant="outline" style={{ fontSize: 12 }} onClick={() => setActiveRun(r.id)}>View</Btn>
                )}
                <Btn variant="ghost" style={{ fontSize: 12, color: C.red }} onClick={() => dispatch({ type: "DELETE_RUN", id: r.id })}>✕</Btn>
              </div>
            </div>
          );
        })}
      </div>

      {showAdd && (
        <AddRunModal state={state} onClose={() => setShowAdd(false)} onSave={(r) => { dispatch({ type: "ADD_RUN", run: r }); setShowAdd(false); setActiveRun(r.id); }} />
      )}
    </div>
  );
}

function RunStatusBadge({ status }) {
  const meta = {
    active: { label: "Active", color: C.blue, bg: C.blueSoft },
    completed: { label: "Completed", color: C.green, bg: C.greenSoft },
    draft: { label: "Draft", color: C.gray500, bg: C.gray100 },
  };
  const m = meta[status] || meta.draft;
  return <span style={{ fontSize: 11, fontWeight: 600, color: m.color, background: m.bg, padding: "2px 8px", borderRadius: 20 }}>{m.label}</span>;
}

function AddRunModal({ state, onClose, onSave }) {
  const { suites } = state;
  const [name, setName] = useState("");
  const [desc, setDesc] = useState("");
  const [selectedSuites, setSelectedSuites] = useState([]);

  const toggle = (id) => setSelectedSuites((p) => p.includes(id) ? p.filter((x) => x !== id) : [...p, id]);

  const tcCount = useMemo(() => {
    const set = new Set();
    selectedSuites.forEach((sid) => {
      const suite = suites.find((s) => s.id === sid);
      if (suite) resolveSuiteTCs(suite, state).forEach((tc) => set.add(tc.id));
    });
    return set.size;
  }, [selectedSuites, state, suites]);

  return (
    <Modal title="New Test Run" onClose={onClose} width={520}>
      <Field label="Run Name" required><Input value={name} onChange={(e) => setName(e.target.value)} placeholder="E.g. Sprint 12 Regression" /></Field>
      <Field label="Description"><Textarea value={desc} onChange={(e) => setDesc(e.target.value)} /></Field>
      <Field label="Select Suites" required>
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
          {suites.map((s) => {
            const checked = selectedSuites.includes(s.id);
            return (
              <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", cursor: "pointer", background: checked ? C.blueSoft : C.white, borderBottom: `1px solid ${C.border}` }}>
                <input type="checkbox" checked={checked} onChange={() => toggle(s.id)} />
                <div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.gray800 }}>{s.name}</div>
                  <div style={{ fontSize: 11, color: C.gray400 }}>{resolveSuiteTCs(s, state).length} test cases</div>
                </div>
              </label>
            );
          })}
        </div>
      </Field>
      {tcCount > 0 && <div style={{ fontSize: 13, color: C.blue, marginBottom: 12 }}>→ {tcCount} unique test cases will be included</div>}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn onClick={onClose} variant="outline">Cancel</Btn>
        <Btn variant="primary" onClick={() => {
          if (!name || selectedSuites.length === 0) return;
          const tcSet = new Set();
          selectedSuites.forEach((sid) => {
            const suite = suites.find((s) => s.id === sid);
            if (suite) resolveSuiteTCs(suite, state).forEach((tc) => tcSet.add(tc.id));
          });
          const cases = [...tcSet].map((id) => {
            const tc = state.testcases.find((t) => t.id === id);
            return { tcId: id, status: "pending", comment: "", executedAt: null };
          });
          onSave({ id: uid(), name, description: desc, suiteIds: selectedSuites, cases, status: "active", createdAt: new Date().toISOString(), completedAt: null });
        }}>Start Run</Btn>
      </div>
    </Modal>
  );
}

// ─── Run Execution View ───────────────────────────────────────────────────────
function RunExecutionView({ run, state, dispatch, onBack }) {
  const [filter, setFilter] = useState("all");
  const [selected, setSelected] = useState(null);

  const getTC = (id) => state.testcases.find((t) => t.id === id);
  const cases = run.cases;
  const filtered = filter === "all" ? cases : cases.filter((c) => c.status === filter);

  const stats = useMemo(() => {
    const total = cases.length;
    const counts = {};
    Object.keys(STATUS_META).forEach((k) => { counts[k] = cases.filter((c) => c.status === k).length; });
    const done = total - counts.pending;
    return { total, ...counts, done };
  }, [cases]);

  const updateCase = (tcId, update) => dispatch({ type: "UPDATE_RUN_CASE", runId: run.id, tcId, update });
  const completeRun = () => dispatch({ type: "COMPLETE_RUN", runId: run.id });

  const allDone = stats.pending === 0;
  const pct = stats.total ? Math.round((stats.pass / stats.total) * 100) : 0;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      {/* Header */}
      <div style={{ padding: "14px 20px", borderBottom: `1px solid ${C.border}`, background: C.white, display: "flex", alignItems: "center", gap: 12 }}>
        <Btn variant="ghost" onClick={onBack} style={{ padding: "4px 8px" }}>← Back</Btn>
        <div style={{ flex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: C.gray800 }}>{run.name}</span>
            <RunStatusBadge status={run.status} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 16, fontSize: 13 }}>
          <span style={{ color: C.green }}>✓ {stats.pass}</span>
          <span style={{ color: C.red }}>✗ {stats.fail}</span>
          <span style={{ color: C.gray400 }}>⋯ {stats.pending}</span>
          <span style={{ fontWeight: 700, color: C.blue }}>{pct}%</span>
        </div>
        {run.status !== "completed" && allDone && (
          <Btn variant="success" onClick={completeRun}>✓ Mark Complete</Btn>
        )}
        {run.status !== "completed" && !allDone && (
          <span style={{ fontSize: 12, color: C.gray400 }}>{stats.done}/{stats.total} done</span>
        )}
      </div>

      {/* Progress bar */}
      <div style={{ height: 4, background: C.gray200 }}>
        <div style={{ height: "100%", background: C.green, width: `${pct}%`, transition: "width 0.3s" }} />
      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* Left: case list */}
        <div style={{ width: 320, flexShrink: 0, borderRight: `1px solid ${C.border}`, overflowY: "auto", display: "flex", flexDirection: "column" }}>
          <div style={{ padding: "10px 12px", borderBottom: `1px solid ${C.border}` }}>
            <Select value={filter} onChange={(e) => setFilter(e.target.value)} style={{ fontSize: 12 }}>
              <option value="all">All ({stats.total})</option>
              <option value="pending">Pending ({stats.pending})</option>
              <option value="pass">Pass ({stats.pass})</option>
              <option value="fail">Fail ({stats.fail})</option>
              <option value="blocked">Blocked ({stats.blocked})</option>
              <option value="skipped">Skipped ({stats.skipped})</option>
            </Select>
          </div>
          {filtered.map((c) => {
            const tc = getTC(c.tcId);
            if (!tc) return null;
            const mod = state.modules.find((m) => m.id === tc.moduleId);
            return (
              <div
                key={c.tcId}
                onClick={() => setSelected(c.tcId)}
                style={{
                  padding: "12px 14px", cursor: "pointer", borderBottom: `1px solid ${C.border}`,
                  background: selected === c.tcId ? C.blueSoft : "transparent",
                  borderLeft: selected === c.tcId ? `3px solid ${C.blue}` : "3px solid transparent",
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 6 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 10, color: C.gray400, fontFamily: "monospace", marginBottom: 3 }}>{tc.tcId}</div>
                    <div style={{ fontSize: 12, fontWeight: 500, color: C.gray800, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tc.title}</div>
                    <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>{mod?.name}</div>
                  </div>
                  <StatusBadge status={c.status} small />
                </div>
              </div>
            );
          })}
        </div>

        {/* Right: case detail + execute */}
        <div style={{ flex: 1, overflowY: "auto" }}>
          {selected ? (
            <CaseExecutor
              runCase={cases.find((c) => c.tcId === selected)}
              tc={getTC(selected)}
              state={state}
              disabled={run.status === "completed"}
              onUpdate={(update) => updateCase(selected, update)}
            />
          ) : (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: C.gray400, fontSize: 14 }}>
              Select a test case to execute
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function CaseExecutor({ runCase, tc, state, onUpdate, disabled }) {
  if (!tc) return null;
  const mod = state.modules.find((m) => m.id === tc.moduleId);
  const proj = state.projects.find((p) => p.id === tc.projectId);

  return (
    <div style={{ padding: 24 }}>
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 11, color: C.gray400, marginBottom: 4 }}>
          {proj?.name} › {mod?.name}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <span style={{ fontFamily: "monospace", fontSize: 12, color: C.blue, background: C.blueSoft, padding: "2px 8px", borderRadius: 6 }}>{tc.tcId}</span>
          <PriorityBadge priority={tc.priority} />
        </div>
        <h3 style={{ margin: "8px 0 4px", fontSize: 17, fontWeight: 700, color: C.gray800 }}>{tc.title}</h3>
        {tc.description && <p style={{ margin: "0 0 16px", color: C.gray500, fontSize: 13 }}>{tc.description}</p>}
      </div>

      <div style={{ background: C.gray50, borderRadius: 8, padding: 14, marginBottom: 14 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: C.gray400, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Test Steps</div>
        <pre style={{ margin: 0, fontSize: 13, color: C.gray700, whiteSpace: "pre-wrap", fontFamily: "inherit", lineHeight: 1.6 }}>{tc.steps}</pre>
      </div>
      <div style={{ background: C.greenSoft, borderRadius: 8, padding: 14, marginBottom: 20 }}>
        <div style={{ fontSize: 11, fontWeight: 700, color: C.green, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Expected Result</div>
        <div style={{ fontSize: 13, color: C.gray700 }}>{tc.expectedResult}</div>
      </div>

      {!disabled && (
        <>
          <div style={{ fontSize: 13, fontWeight: 600, color: C.gray700, marginBottom: 10 }}>Update Result</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
            {Object.entries(STATUS_META).filter(([k]) => k !== "in-progress").map(([k, m]) => (
              <Btn
                key={k}
                onClick={() => onUpdate({ status: k, executedAt: new Date().toISOString() })}
                variant="outline"
                style={{
                  background: runCase.status === k ? m.bg : C.white,
                  borderColor: runCase.status === k ? m.color : C.border,
                  color: runCase.status === k ? m.color : C.gray600,
                  fontWeight: runCase.status === k ? 700 : 500,
                  padding: "7px 16px",
                }}
              >
                {m.label}
              </Btn>
            ))}
          </div>
          <Field label="Comment / Notes">
            <Textarea value={runCase.comment} onChange={(e) => onUpdate({ comment: e.target.value })} placeholder="Optional: describe what you observed..." />
          </Field>
        </>
      )}

      {disabled && (
        <div style={{ background: C.gray50, borderRadius: 8, padding: 14 }}>
          <div style={{ fontSize: 12, color: C.gray500, marginBottom: 4 }}>Result</div>
          <StatusBadge status={runCase.status} />
          {runCase.comment && <div style={{ fontSize: 13, color: C.gray600, marginTop: 8 }}>{runCase.comment}</div>}
        </div>
      )}
    </div>
  );
}

// ─── Reports View ─────────────────────────────────────────────────────────────
function ReportsView({ state }) {
  const completedRuns = state.runs.filter((r) => r.status === "completed");
  const [selectedRun, setSelectedRun] = useState(completedRuns[0]?.id || null);
  const run = state.runs.find((r) => r.id === selectedRun);

  if (completedRuns.length === 0) {
    return (
      <div style={{ padding: 40, textAlign: "center", color: C.gray400 }}>
        <div style={{ fontSize: 40, marginBottom: 12 }}>📊</div>
        <div style={{ fontSize: 15, fontWeight: 500, color: C.gray600 }}>No completed runs yet</div>
        <div style={{ fontSize: 13, marginTop: 4 }}>Complete a test run to see reports</div>
      </div>
    );
  }

  return (
    <div style={{ padding: 24, maxWidth: 1000, margin: "0 auto" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: C.gray800 }}>Reports</h2>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <Select value={selectedRun || ""} onChange={(e) => setSelectedRun(e.target.value)} style={{ width: 240 }}>
            {completedRuns.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </Select>
          <Btn variant="outline" onClick={() => exportReportPDF(run, state)} style={{ fontSize: 12 }}>↓ PDF</Btn>
          <Btn variant="outline" onClick={() => exportReportImage()} style={{ fontSize: 12 }}>↓ Image</Btn>
        </div>
      </div>

      {run && <ReportContent run={run} state={state} />}
    </div>
  );
}

function ReportContent({ run, state }) {
  const cases = run.cases;
  const total = cases.length;
  const byStatus = {};
  Object.keys(STATUS_META).forEach((k) => { byStatus[k] = cases.filter((c) => c.status === k).length; });
  const passed = byStatus.pass;
  const failed = byStatus.fail;
  const passRate = total ? Math.round((passed / total) * 100) : 0;

  const byProject = {};
  cases.forEach((c) => {
    const tc = state.testcases.find((t) => t.id === c.tcId);
    if (!tc) return;
    const proj = state.projects.find((p) => p.id === tc.projectId);
    const key = proj?.name || "Unknown";
    if (!byProject[key]) byProject[key] = { pass: 0, fail: 0, total: 0 };
    byProject[key].total++;
    byProject[key][c.status === "pass" ? "pass" : c.status === "fail" ? "fail" : "other"] = (byProject[key][c.status === "pass" ? "pass" : c.status === "fail" ? "fail" : "other"] || 0) + 1;
  });

  const byModule = {};
  cases.forEach((c) => {
    const tc = state.testcases.find((t) => t.id === c.tcId);
    if (!tc) return;
    const mod = state.modules.find((m) => m.id === tc.moduleId);
    const key = mod?.name || "Unknown";
    if (!byModule[key]) byModule[key] = { pass: 0, fail: 0, total: 0 };
    byModule[key].total++;
    if (c.status === "pass") byModule[key].pass++;
    else if (c.status === "fail") byModule[key].fail++;
  });

  const byPriority = {};
  ["high", "medium", "low"].forEach((p) => {
    const relevant = cases.filter((c) => state.testcases.find((t) => t.id === c.tcId)?.priority === p);
    byPriority[p] = { total: relevant.length, pass: relevant.filter((c) => c.status === "pass").length, fail: relevant.filter((c) => c.status === "fail").length };
  });

  return (
    <div id="report-content">
      {/* Header card */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24, marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <h3 style={{ margin: "0 0 4px", fontSize: 18, fontWeight: 700, color: C.gray800 }}>{run.name}</h3>
            <div style={{ fontSize: 13, color: C.gray400 }}>
              Completed {run.completedAt ? new Date(run.completedAt).toLocaleString() : "—"}
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 36, fontWeight: 800, color: passRate >= 80 ? C.green : passRate >= 60 ? C.amber : C.red }}>{passRate}%</div>
            <div style={{ fontSize: 12, color: C.gray400 }}>Pass Rate</div>
          </div>
        </div>

        {/* Metric cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 10, marginTop: 20 }}>
          {[
            { label: "Total", val: total, color: C.gray700 },
            { label: "Passed", val: byStatus.pass, color: C.green },
            { label: "Failed", val: byStatus.fail, color: C.red },
            { label: "Blocked", val: byStatus.blocked, color: C.amber },
            { label: "Skipped", val: byStatus.skipped, color: C.purple },
            { label: "Pending", val: byStatus.pending, color: C.gray400 },
          ].map((m) => (
            <div key={m.label} style={{ background: C.gray50, borderRadius: 10, padding: "12px 14px", textAlign: "center" }}>
              <div style={{ fontSize: 24, fontWeight: 800, color: m.color }}>{m.val}</div>
              <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>{m.label}</div>
            </div>
          ))}
        </div>

        {/* Progress bar */}
        <div style={{ marginTop: 16 }}>
          <div style={{ height: 8, background: C.gray100, borderRadius: 8, overflow: "hidden", display: "flex" }}>
            {passed > 0 && <div style={{ width: `${(passed / total) * 100}%`, background: C.green }} />}
            {failed > 0 && <div style={{ width: `${(failed / total) * 100}%`, background: C.red }} />}
            {byStatus.blocked > 0 && <div style={{ width: `${(byStatus.blocked / total) * 100}%`, background: C.amber }} />}
            {byStatus.skipped > 0 && <div style={{ width: `${(byStatus.skipped / total) * 100}%`, background: C.purple }} />}
          </div>
          <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 11 }}>
            {[["Pass", C.green], ["Fail", C.red], ["Blocked", C.amber], ["Skipped", C.purple]].map(([l, c]) => (
              <span key={l} style={{ display: "flex", alignItems: "center", gap: 4, color: C.gray400 }}>
                <span style={{ width: 8, height: 8, borderRadius: 2, background: c, display: "inline-block" }} />{l}
              </span>
            ))}
          </div>
        </div>
      </div>

      {/* Charts row */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        {/* By Module */}
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20 }}>
          <h4 style={{ margin: "0 0 16px", fontSize: 14, fontWeight: 600, color: C.gray700 }}>Results by Module</h4>
          {Object.entries(byModule).map(([name, data]) => {
            const p = data.total ? Math.round((data.pass / data.total) * 100) : 0;
            return (
              <div key={name} style={{ marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.gray600, marginBottom: 4 }}>
                  <span>{name}</span>
                  <span style={{ fontWeight: 600, color: p >= 80 ? C.green : p >= 60 ? C.amber : C.red }}>{p}%</span>
                </div>
                <div style={{ height: 6, background: C.gray100, borderRadius: 4, overflow: "hidden", display: "flex" }}>
                  <div style={{ width: `${p}%`, background: p >= 80 ? C.green : p >= 60 ? C.amber : C.red, borderRadius: 4 }} />
                </div>
                <div style={{ fontSize: 11, color: C.gray400, marginTop: 2 }}>✓{data.pass} ✗{data.fail} of {data.total}</div>
              </div>
            );
          })}
        </div>

        {/* By Priority */}
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20 }}>
          <h4 style={{ margin: "0 0 16px", fontSize: 14, fontWeight: 600, color: C.gray700 }}>Results by Priority</h4>
          {[["high", "High", C.red], ["medium", "Medium", C.amber], ["low", "Low", C.green]].map(([key, label, color]) => {
            const data = byPriority[key];
            const p = data.total ? Math.round((data.pass / data.total) * 100) : 0;
            return (
              <div key={key} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ width: 8, height: 8, borderRadius: 2, background: color, display: "inline-block" }} />
                    <span style={{ fontSize: 12, color: C.gray600 }}>{label}</span>
                  </div>
                  <span style={{ fontSize: 12, fontWeight: 600, color: C.gray600 }}>{data.total} cases</span>
                </div>
                <div style={{ height: 20, background: C.gray100, borderRadius: 4, overflow: "hidden", display: "flex" }}>
                  {data.pass > 0 && <div style={{ width: `${(data.pass / data.total) * 100}%`, background: C.green, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {data.pass > 0 && <span style={{ fontSize: 10, color: C.white, fontWeight: 700 }}>{data.pass}</span>}
                  </div>}
                  {data.fail > 0 && <div style={{ width: `${(data.fail / data.total) * 100}%`, background: C.red, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <span style={{ fontSize: 10, color: C.white, fontWeight: 700 }}>{data.fail}</span>
                  </div>}
                </div>
              </div>
            );
          })}

          {/* Donut-style stat */}
          <div style={{ marginTop: 20, textAlign: "center" }}>
            <svg width={160} height={160} viewBox="0 0 160 160" style={{ display: "block", margin: "0 auto" }}>
              <DonutChart passed={passed} failed={failed} other={total - passed - failed} total={total} />
            </svg>
            <div style={{ fontSize: 12, color: C.gray400, marginTop: 4 }}>Overall distribution</div>
          </div>
        </div>
      </div>

      {/* Failed cases detail */}
      {failed > 0 && (
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20, marginBottom: 16 }}>
          <h4 style={{ margin: "0 0 14px", fontSize: 14, fontWeight: 600, color: C.red }}>Failed Test Cases ({failed})</h4>
          {cases.filter((c) => c.status === "fail").map((c) => {
            const tc = state.testcases.find((t) => t.id === c.tcId);
            if (!tc) return null;
            const mod = state.modules.find((m) => m.id === tc.moduleId);
            return (
              <div key={c.tcId} style={{ borderLeft: `3px solid ${C.red}`, paddingLeft: 12, marginBottom: 12, paddingBottom: 12, borderBottom: `1px solid ${C.border}` }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 4 }}>
                  <span style={{ fontSize: 11, fontFamily: "monospace", color: C.blue, background: C.blueSoft, padding: "1px 7px", borderRadius: 4 }}>{tc.tcId}</span>
                  <PriorityBadge priority={tc.priority} />
                  <span style={{ fontSize: 11, color: C.gray400 }}>{mod?.name}</span>
                </div>
                <div style={{ fontSize: 13, fontWeight: 500, color: C.gray800 }}>{tc.title}</div>
                {c.comment && <div style={{ fontSize: 12, color: C.gray500, marginTop: 4, fontStyle: "italic" }}>"{c.comment}"</div>}
              </div>
            );
          })}
        </div>
      )}

      {/* Full case table */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20 }}>
        <h4 style={{ margin: "0 0 14px", fontSize: 14, fontWeight: 600, color: C.gray700 }}>All Test Cases</h4>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: C.gray50 }}>
              {["TC ID", "Title", "Module", "Priority", "Status", "Comment"].map((h) => (
                <th key={h} style={{ padding: "8px 12px", textAlign: "left", color: C.gray500, fontWeight: 600, borderBottom: `1px solid ${C.border}`, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.4 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {cases.map((c) => {
              const tc = state.testcases.find((t) => t.id === c.tcId);
              if (!tc) return null;
              const mod = state.modules.find((m) => m.id === tc.moduleId);
              return (
                <tr key={c.tcId} style={{ borderBottom: `1px solid ${C.border}` }}>
                  <td style={{ padding: "8px 12px", fontFamily: "monospace", color: C.blue, fontSize: 11 }}>{tc.tcId}</td>
                  <td style={{ padding: "8px 12px", color: C.gray800, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{tc.title}</td>
                  <td style={{ padding: "8px 12px", color: C.gray500 }}>{mod?.name}</td>
                  <td style={{ padding: "8px 12px" }}><PriorityBadge priority={tc.priority} /></td>
                  <td style={{ padding: "8px 12px" }}><StatusBadge status={c.status} small /></td>
                  <td style={{ padding: "8px 12px", color: C.gray400, fontStyle: "italic", maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.comment || "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DonutChart({ passed, failed, other, total }) {
  if (total === 0) return null;
  const cx = 80, cy = 80, r = 60, sw = 20;
  const circ = 2 * Math.PI * r;
  const passAngle = (passed / total) * 360;
  const failAngle = (failed / total) * 360;
  const otherAngle = (other / total) * 360;
  const segs = [
    { angle: passAngle, color: C.green },
    { angle: failAngle, color: C.red },
    { angle: otherAngle, color: C.gray300 },
  ];
  let cumAngle = -90;
  const arcs = segs.map((s) => {
    const start = cumAngle;
    cumAngle += s.angle;
    return { ...s, start };
  });
  const polarToXY = (angle, radius) => {
    const rad = (angle * Math.PI) / 180;
    return { x: cx + radius * Math.cos(rad), y: cy + radius * Math.sin(rad) };
  };
  return (
    <>
      {arcs.map((arc, i) => {
        if (arc.angle === 0) return null;
        const s = polarToXY(arc.start, r);
        const e = polarToXY(arc.start + arc.angle, r);
        const large = arc.angle > 180 ? 1 : 0;
        return (
          <path key={i} d={`M ${s.x} ${s.y} A ${r} ${r} 0 ${large} 1 ${e.x} ${e.y}`} fill="none" stroke={arc.color} strokeWidth={sw} />
        );
      })}
      <text x={cx} y={cy - 6} textAnchor="middle" fontSize={22} fontWeight={800} fill={C.gray800}>{Math.round((passed / total) * 100)}%</text>
      <text x={cx} y={cy + 14} textAnchor="middle" fontSize={11} fill={C.gray400}>pass rate</text>
    </>
  );
}

function exportReportPDF(run, state) {
  window.print();
}

function exportReportImage() {
  alert("To save as image: use your browser's screenshot tool or right-click the report and select 'Save as image'.");
}

// ─── Excel Import / Export ────────────────────────────────────────────────────
function exportToExcel(state) {
  const wb = XLSX.utils.book_new();

  const projSheet = XLSX.utils.json_to_sheet(state.projects.map((p) => ({
    ID: p.id, Name: p.name, Key: p.key, Description: p.description, CreatedAt: p.createdAt,
  })));
  XLSX.utils.book_append_sheet(wb, projSheet, "Projects");

  const modSheet = XLSX.utils.json_to_sheet(state.modules.map((m) => ({
    ID: m.id, ProjectID: m.projectId, Name: m.name, Key: m.key, Description: m.description,
  })));
  XLSX.utils.book_append_sheet(wb, modSheet, "Modules");

  const tcSheet = XLSX.utils.json_to_sheet(state.testcases.map((tc) => ({
    ID: tc.id, ProjectID: tc.projectId, ModuleID: tc.moduleId, TCID: tc.tcId, Num: tc.num,
    Title: tc.title, Description: tc.description, Steps: tc.steps, ExpectedResult: tc.expectedResult,
    Priority: tc.priority, Tags: tc.tags, Status: tc.status,
  })));
  XLSX.utils.book_append_sheet(wb, tcSheet, "TestCases");

  const suiteSheet = XLSX.utils.json_to_sheet(state.suites.map((s) => ({
    ID: s.id, Name: s.name, Description: s.description, CreatedAt: s.createdAt,
    Inclusions: JSON.stringify(s.inclusions),
  })));
  XLSX.utils.book_append_sheet(wb, suiteSheet, "Suites");

  const runRows = [];
  state.runs.forEach((r) => {
    r.cases.forEach((c) => {
      runRows.push({ RunID: r.id, RunName: r.name, RunStatus: r.status, CreatedAt: r.createdAt, CompletedAt: r.completedAt || "", TCID: c.tcId, CaseStatus: c.status, Comment: c.comment, ExecutedAt: c.executedAt || "" });
    });
  });
  const runSheet = XLSX.utils.json_to_sheet(runRows);
  XLSX.utils.book_append_sheet(wb, runSheet, "Runs");

  XLSX.writeFile(wb, "TestManager_Export.xlsx");
}

function importFromExcel(file, dispatch) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const getSheet = (name) => {
        const ws = wb.Sheets[name];
        return ws ? XLSX.utils.sheet_to_json(ws) : [];
      };

      const projects = getSheet("Projects").map((r) => ({ id: r.ID || uid(), name: r.Name, key: r.Key, description: r.Description || "", createdAt: r.CreatedAt || new Date().toISOString() }));
      const modules = getSheet("Modules").map((r) => ({ id: r.ID || uid(), projectId: r.ProjectID, name: r.Name, key: r.Key, description: r.Description || "" }));
      const testcases = getSheet("TestCases").map((r) => ({ id: r.ID || uid(), projectId: r.ProjectID, moduleId: r.ModuleID, tcId: r.TCID, num: r.Num || 1, title: r.Title, description: r.Description || "", steps: r.Steps || "", expectedResult: r.ExpectedResult || "", priority: r.Priority || "medium", tags: r.Tags || "", status: r.Status || "active" }));
      const suites = getSheet("Suites").map((r) => ({ id: r.ID || uid(), name: r.Name, description: r.Description || "", createdAt: r.CreatedAt || new Date().toISOString(), inclusions: (() => { try { return JSON.parse(r.Inclusions || "[]"); } catch { return []; } })() }));

      const runMap = {};
      getSheet("Runs").forEach((r) => {
        if (!runMap[r.RunID]) {
          runMap[r.RunID] = { id: r.RunID, name: r.RunName, description: "", suiteIds: [], cases: [], status: r.RunStatus || "active", createdAt: r.CreatedAt || new Date().toISOString(), completedAt: r.CompletedAt || null };
        }
        runMap[r.RunID].cases.push({ tcId: r.TCID, status: r.CaseStatus || "pending", comment: r.Comment || "", executedAt: r.ExecutedAt || null });
      });
      const runs = Object.values(runMap);

      dispatch({ type: "IMPORT", data: { projects, modules, testcases, suites, runs } });
      alert("Import successful!");
    } catch (err) {
      alert("Import failed: " + err.message);
    }
  };
  reader.readAsBinaryString(file);
}

// ─── Reducer ──────────────────────────────────────────────────────────────────
function reducer(state, action) {
  switch (action.type) {
    case "ADD_PROJECT": return { ...state, projects: [...state.projects, action.project] };
    case "ADD_MODULE": return { ...state, modules: [...state.modules, action.module] };
    case "ADD_TC": return { ...state, testcases: [...state.testcases, action.tc] };
    case "EDIT_TC": return { ...state, testcases: state.testcases.map((tc) => tc.id === action.tc.id ? action.tc : tc) };
    case "DELETE_TC": return { ...state, testcases: state.testcases.filter((tc) => tc.id !== action.id) };
    case "ADD_SUITE": return { ...state, suites: [...state.suites, action.suite] };
    case "EDIT_SUITE": return { ...state, suites: state.suites.map((s) => s.id === action.suite.id ? action.suite : s) };
    case "DELETE_SUITE": return { ...state, suites: state.suites.filter((s) => s.id !== action.id) };
    case "ADD_RUN": return { ...state, runs: [...state.runs, action.run] };
    case "DELETE_RUN": return { ...state, runs: state.runs.filter((r) => r.id !== action.id) };
    case "UPDATE_RUN_CASE": return {
      ...state,
      runs: state.runs.map((r) => r.id === action.runId
        ? { ...r, cases: r.cases.map((c) => c.tcId === action.tcId ? { ...c, ...action.update } : c) }
        : r),
    };
    case "COMPLETE_RUN": return {
      ...state,
      runs: state.runs.map((r) => r.id === action.runId ? { ...r, status: "completed", completedAt: new Date().toISOString() } : r),
    };
    case "IMPORT": return { ...state, ...action.data };
    default: return state;
  }
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [state, dispatch] = useState(() => makeInitialState());
  const [nav, setNav] = useState("projects");
  const fileRef = useRef();

  const dispatchFn = useCallback((action) => {
    setState((s) => reducer(s, action));
  }, []);

  function setState(fn) {
    dispatch((prev) => typeof fn === "function" ? fn(prev) : fn);
  }

  const totalTCs = state.testcases.length;
  const totalRuns = state.runs.length;
  const completedRuns = state.runs.filter((r) => r.status === "completed").length;

  return (
    <div style={{ fontFamily: "'DM Sans', 'Segoe UI', system-ui, sans-serif", height: "100vh", display: "flex", flexDirection: "column", background: C.gray50, color: C.gray800 }}>
      {/* Top bar */}
      <div style={{
        background: C.white, borderBottom: `1px solid ${C.border}`,
        padding: "0 20px", height: 56, display: "flex", alignItems: "center", gap: 0, flexShrink: 0,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginRight: 32 }}>
          <div
  style={{
    width: 28,
    height: 28,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    overflow: "hidden",
  }}
>
  <img
    src={logo1}
    alt="Logo"
    style={{
      maxWidth: "100%",
      maxHeight: "100%",
      objectFit: "contain",
    }}
  />
</div>
          <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.1 }}>
  <span
    style={{
      fontSize: 16,
      fontWeight: 800,
      color: C.gray800,
      letterSpacing: 0.3,
    }}
  >
    TMT
  </span>

  <span
    style={{
      fontSize: 11,
      fontWeight: 500,
      color: C.gray500,
      marginTop: 2,
    }}
  >
    V1.1.113{" "}
    <span
      style={{
        color: C.blue,
        fontWeight: 700,
      }}
    >
    </span>
  </span>
</div>
        </div>

        <nav style={{ display: "flex", gap: 2, flex: 1 }}>
          {NAV_ITEMS.map((item) => (
            <button
              key={item.id}
              onClick={() => setNav(item.id)}
              style={{
                background: nav === item.id ? C.blueSoft : "transparent",
                border: "none", borderRadius: 8,
                padding: "6px 14px", cursor: "pointer", fontFamily: "inherit",
                fontSize: 13, fontWeight: nav === item.id ? 600 : 500,
                color: nav === item.id ? C.blue : C.gray500,
                display: "flex", alignItems: "center", gap: 6,
              }}
            >
              <span>{item.icon}</span> {item.label}
            </button>
          ))}
        </nav>

        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <div style={{ display: "flex", gap: 12, fontSize: 12, color: C.gray400, marginRight: 12 }}>
            <span>{state.projects.length} projects</span>
            <span>{totalTCs} tests</span>
            <span>{completedRuns}/{totalRuns} runs done</span>
          </div>
          <Btn variant="outline" style={{ fontSize: 12 }} onClick={() => exportToExcel(state)}>↑ Export</Btn>
          <Btn variant="outline" style={{ fontSize: 12 }} onClick={() => fileRef.current.click()}>↓ Import</Btn>
          <input ref={fileRef} type="file" accept=".xlsx" style={{ display: "none" }} onChange={(e) => { if (e.target.files[0]) importFromExcel(e.target.files[0], dispatchFn); e.target.value = ""; }} />
        </div>
      </div>

      {/* Main content */}
      <div style={{ flex: 1, overflow: "hidden" }}>
        {nav === "projects" && <ProjectsView state={state} dispatch={dispatchFn} />}
        {nav === "suites" && <SuitesView state={state} dispatch={dispatchFn} />}
        {nav === "runs" && <RunsView state={state} dispatch={dispatchFn} />}
        {nav === "reports" && <ReportsView state={state} />}
      </div>
    </div>
  );
}