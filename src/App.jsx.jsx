import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase.js";

// ─── SUPABASE HOOKS ──────────────────────────────────────────────────────────
function useProjects() {
  const [projects, setProjects] = useState([]);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    const { data } = await supabase.from("projects").select("*").order("created_at", { ascending: false });
    if (data) setProjects(data);
    setLoading(false);
  }, []);

  useEffect(() => {
    load();
    // Realtime: обновления у всей команды сразу
    const channel = supabase.channel("projects_rt")
      .on("postgres_changes", { event: "*", schema: "public", table: "projects" }, load)
      .subscribe();
    return () => supabase.removeChannel(channel);
  }, [load]);

  const addProject = async (f) => {
    const { data } = await supabase.from("projects").insert([{ ...f, amount: Number(f.amount) }]).select().single();
    if (data) setProjects(p => [data, ...p]);
  };

  const updateProject = async (id, f) => {
    const { id: _, created_at, updated_at, ...updateData } = f;
    const { data } = await supabase.from("projects").update({ ...updateData, amount: Number(f.amount) }).eq("id", id).select().single();
    if (data) setProjects(p => p.map(x => x.id === id ? data : x));
  };

  const deleteProject = async (id) => {
    await supabase.from("projects").delete().eq("id", id);
    setProjects(p => p.filter(x => x.id !== id));
  };

  return { projects, setProjects, loading, addProject, updateProject, deleteProject };
}

function useTasks() {
  const [tasks, setTasks] = useState([]);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    const { data } = await supabase.from("tasks").select("*").order("created_at", { ascending: false });
    if (data) setTasks(data);
    setLoading(false);
  }, []);

  useEffect(() => {
    load();
    const channel = supabase.channel("tasks_rt")
      .on("postgres_changes", { event: "*", schema: "public", table: "tasks" }, load)
      .subscribe();
    return () => supabase.removeChannel(channel);
  }, [load]);

  const addTask = async (f) => {
    const { data } = await supabase.from("tasks").insert([f]).select().single();
    if (data) setTasks(t => [data, ...t]);
  };

  const updateTask = async (id, f) => {
    const { id: _, created_at, updated_at, ...updateData } = f;
    const { data } = await supabase.from("tasks").update(updateData).eq("id", id).select().single();
    if (data) setTasks(t => t.map(x => x.id === id ? data : x));
  };

  const deleteTask = async (id) => {
    await supabase.from("tasks").delete().eq("id", id);
    setTasks(t => t.filter(x => x.id !== id));
  };

  return { tasks, setTasks, loading, addTask, updateTask, deleteTask };
}

function useManagers() {
  const [managers, setManagers] = useState([]);

  const load = useCallback(async () => {
    const { data } = await supabase.from("managers").select("*").eq("active", true).order("name");
    if (data) setManagers(data.map(m => m.name));
  }, []);

  useEffect(() => {
    load();
    const channel = supabase.channel("managers_rt")
      .on("postgres_changes", { event: "*", schema: "public", table: "managers" }, load)
      .subscribe();
    return () => supabase.removeChannel(channel);
  }, [load]);

  const addManager = async (name, role = "Менеджер") => {
    await supabase.from("managers").insert([{ name, role }]);
    await load();
  };

  const removeManager = async (name) => {
    await supabase.from("managers").update({ active: false }).eq("name", name);
    await load();
  };

  return { managers, addManager, removeManager };
}


const exportToExcel = (projects, tasks, sheetName = "all") => {
  const wb = XLSX.utils.book_new();

  if (sheetName === "all" || sheetName === "projects") {
    const projRows = projects.map(p => ({
      "Клиент": p.client,
      "Сумма (₸)": Number(p.amount),
      "Тип финансирования": p.type,
      "Стадия": p.stage,
      "Менеджер": p.manager,
      "Дата поступления": p.date,
      "Примечания": p.notes || "",
    }));
    const wsP = XLSX.utils.json_to_sheet(projRows);
    wsP["!cols"] = [{ wch: 30 }, { wch: 16 }, { wch: 20 }, { wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, wsP, "Проекты");
  }

  if (sheetName === "all" || sheetName === "tasks") {
    const taskRows = tasks.map(t => ({
      "Тип задачи": t.type,
      "Клиент": t.client,
      "Менеджер": t.manager,
      "Статус": t.status,
      "Приоритет": t.priority,
      "Срок": t.deadline || "",
      "Просрочено": isOverdue(t.deadline, t.status) ? "Да" : "Нет",
      "Примечания": t.notes || "",
    }));
    const wsT = XLSX.utils.json_to_sheet(taskRows);
    wsT["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 18 }, { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, wsT, "Задачи");
  }

  if (sheetName === "all") {
    // Summary sheet — derive managers from data
    const managerNames = [...new Set([...projects.map(p => p.manager), ...tasks.map(t => t.manager)])].filter(Boolean);
    const summary = managerNames.map(m => ({
      "Менеджер": m,
      "Активных проектов": projects.filter(p => p.manager === m && p.stage !== "Выдача").length,
      "Активных задач": tasks.filter(t => t.manager === m && t.status !== "Выполнено").length,
      "Просроченных задач": tasks.filter(t => t.manager === m && isOverdue(t.deadline, t.status)).length,
      "Выполнено задач": tasks.filter(t => t.manager === m && t.status === "Выполнено").length,
    }));
    const wsS = XLSX.utils.json_to_sheet(summary);
    wsS["!cols"] = [{ wch: 20 }, { wch: 20 }, { wch: 18 }, { wch: 22 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsS, "Команда");
  }

  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `МСБ_Отчёт_${date}.xlsx`);
};

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
const PROJECT_STAGES = ["Рассмотрение", "Анализ", "Кредитный комитет", "Решение", "Выдача"];
const TASK_TYPES = ["Реструктуризация", "Высвобождение залога", "Предоставление", "Письмо клиента", "Мониторинг"];
const TASK_STATUSES = ["Новая", "В работе", "На согласовании", "Выполнено"];
const PRIORITIES = ["Высокий", "Средний", "Низкий"];
const today = new Date();

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const fmt = (n) => new Intl.NumberFormat("ru-KZ", { maximumFractionDigits: 0 }).format(n) + " ₸";
const fmtM = (n) => (n >= 1e6 ? (n / 1e6).toFixed(1) + " млн ₸" : fmt(n));
const isOverdue = (d, status) => status !== "Выполнено" && new Date(d) < today;
const daysLeft = (d) => Math.ceil((new Date(d) - today) / 86400000);

// ─── COLORS ──────────────────────────────────────────────────────────────────
const C = {
  bg: "#07111E",
  surface: "#0E1F33",
  card: "#122540",
  border: "#1C3454",
  accent: "#E8A020",
  accentDim: "rgba(232,160,32,0.12)",
  text: "#DCE8F5",
  muted: "#5E7A96",
  success: "#22C997",
  successDim: "rgba(34,201,151,0.12)",
  warning: "#F59E0B",
  warningDim: "rgba(245,158,11,0.12)",
  danger: "#F06969",
  dangerDim: "rgba(240,105,105,0.12)",
  info: "#60A5FA",
  infoDim: "rgba(96,165,250,0.12)",
};

const stageColor = (s) => ({
  "Рассмотрение": C.muted,
  "Анализ": C.info,
  "Кредитный комитет": C.warning,
  "Решение": C.accent,
  "Выдача": C.success,
}[s] || C.muted);

const statusColor = (s) => ({
  "Новая": C.muted,
  "В работе": C.info,
  "На согласовании": C.warning,
  "Выполнено": C.success,
}[s] || C.muted);

const priorityColor = (p) => ({ "Высокий": C.danger, "Средний": C.warning, "Низкий": C.success }[p] || C.muted);

const typeIcon = (t) => ({ "Реструктуризация": "🔄", "Высвобождение залога": "🔓", "Предоставление": "📤", "Письмо клиента": "✉️", "Мониторинг": "🔍" }[t] || "📌");

// ─── UI PRIMITIVES ────────────────────────────────────────────────────────────
const Badge = ({ color, bg, children, small }) => (
  <span style={{
    display: "inline-flex", alignItems: "center", gap: 4,
    background: bg || "rgba(255,255,255,0.07)", color: color || C.text,
    borderRadius: 6, padding: small ? "2px 7px" : "3px 10px",
    fontSize: small ? 11 : 12, fontWeight: 600, whiteSpace: "nowrap",
    border: `1px solid ${color || C.border}33`,
  }}>{children}</span>
);

const Chip = ({ label, color }) => (
  <span style={{
    display: "inline-block", padding: "2px 8px", borderRadius: 20,
    fontSize: 11, fontWeight: 700, letterSpacing: 0.3,
    background: color + "20", color: color, border: `1px solid ${color}44`,
  }}>{label}</span>
);

const StatCard = ({ icon, label, value, sub, accent, onClick }) => (
  <div onClick={onClick} style={{
    background: C.card, border: `1px solid ${C.border}`, borderRadius: 14,
    padding: "18px 20px", cursor: onClick ? "pointer" : "default",
    transition: "transform .15s, box-shadow .15s",
    position: "relative", overflow: "hidden",
  }}
    onMouseEnter={e => { if (onClick) { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = `0 8px 24px rgba(0,0,0,.4)`; } }}
    onMouseLeave={e => { e.currentTarget.style.transform = ""; e.currentTarget.style.boxShadow = ""; }}
  >
    <div style={{ position: "absolute", top: 0, right: 0, width: 80, height: 80, background: (accent || C.accent) + "08", borderRadius: "0 14px 0 80px" }} />
    <div style={{ fontSize: 22, marginBottom: 8 }}>{icon}</div>
    <div style={{ fontSize: 26, fontWeight: 800, color: accent || C.text, fontFamily: "'DM Serif Display', serif", lineHeight: 1 }}>{value}</div>
    <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{label}</div>
    {sub && <div style={{ fontSize: 11, color: accent || C.accent, marginTop: 6, fontWeight: 600 }}>{sub}</div>}
  </div>
);

const Modal = ({ open, onClose, title, children }) => {
  if (!open) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.7)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}
      onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, width: "100%", maxWidth: 540, maxHeight: "85vh", overflow: "auto", boxShadow: "0 24px 64px rgba(0,0,0,.6)" }}>
        <div style={{ padding: "20px 24px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text, fontFamily: "'DM Serif Display', serif" }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: C.muted, fontSize: 20, cursor: "pointer", lineHeight: 1 }}>×</button>
        </div>
        <div style={{ padding: "16px 24px 24px" }}>{children}</div>
      </div>
    </div>
  );
};

const Field = ({ label, children }) => (
  <div style={{ marginBottom: 14 }}>
    <label style={{ display: "block", fontSize: 11, color: C.muted, fontWeight: 600, marginBottom: 5, textTransform: "uppercase", letterSpacing: 0.5 }}>{label}</label>
    {children}
  </div>
);

const Input = ({ value, onChange, placeholder, type = "text" }) => (
  <input type={type} value={value} onChange={onChange} placeholder={placeholder}
    style={{ width: "100%", background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "9px 12px", color: C.text, fontSize: 13, outline: "none", boxSizing: "border-box" }} />
);

const Select = ({ value, onChange, options }) => (
  <select value={value} onChange={onChange}
    style={{ width: "100%", background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "9px 12px", color: C.text, fontSize: 13, outline: "none" }}>
    {options.map(o => <option key={o} value={o}>{o}</option>)}
  </select>
);

const Textarea = ({ value, onChange, placeholder }) => (
  <textarea value={value} onChange={onChange} placeholder={placeholder} rows={3}
    style={{ width: "100%", background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "9px 12px", color: C.text, fontSize: 13, outline: "none", resize: "vertical", boxSizing: "border-box", fontFamily: "inherit" }} />
);

const Btn = ({ onClick, children, variant = "primary", small }) => {
  const styles = {
    primary: { background: C.accent, color: "#0A0A0A" },
    ghost: { background: "transparent", color: C.muted, border: `1px solid ${C.border}` },
    danger: { background: C.danger + "22", color: C.danger, border: `1px solid ${C.danger}44` },
  };
  return (
    <button onClick={onClick} style={{
      ...styles[variant], borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 700,
      padding: small ? "6px 12px" : "9px 18px", fontSize: small ? 12 : 13,
      transition: "opacity .15s",
    }}
      onMouseEnter={e => e.currentTarget.style.opacity = ".85"}
      onMouseLeave={e => e.currentTarget.style.opacity = "1"}
    >{children}</button>
  );
};

// ─── FORMS ────────────────────────────────────────────────────────────────────
const blankTask = () => ({ type: TASK_TYPES[0], client: "", manager: "", status: "Новая", priority: "Средний", deadline: "", notes: "" });
const blankProject = () => ({ client: "", amount: "", manager: "", stage: "Рассмотрение", type: "Оборотный", date: new Date().toISOString().slice(0, 10), notes: "" });

function TaskForm({ init, onSave, onClose, managers = [] }) {
  const [f, setF] = useState(init || { type: TASK_TYPES[0], client: "", manager: managers[0] || "", status: "Новая", priority: "Средний", deadline: "", notes: "" });
  const set = (k) => (e) => setF(p => ({ ...p, [k]: e.target.value }));
  return (
    <div>
      <Field label="Тип задачи"><Select value={f.type} onChange={set("type")} options={TASK_TYPES} /></Field>
      <Field label="Клиент"><Input value={f.client} onChange={set("client")} placeholder="Наименование клиента" /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="Ответственный"><Select value={f.manager} onChange={set("manager")} options={managers} /></Field>
        <Field label="Статус"><Select value={f.status} onChange={set("status")} options={TASK_STATUSES} /></Field>
        <Field label="Приоритет"><Select value={f.priority} onChange={set("priority")} options={PRIORITIES} /></Field>
        <Field label="Срок исполнения"><Input type="date" value={f.deadline} onChange={set("deadline")} /></Field>
      </div>
      <Field label="Примечания"><Textarea value={f.notes} onChange={set("notes")} placeholder="Дополнительная информация..." /></Field>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <Btn variant="ghost" onClick={onClose}>Отмена</Btn>
        <Btn onClick={() => { if (f.client) { onSave(f); onClose(); } }}>Сохранить</Btn>
      </div>
    </div>
  );
}

function ProjectForm({ init, onSave, onClose, managers = [] }) {
  const [f, setF] = useState(init || { client: "", amount: "", manager: managers[0] || "", stage: "Рассмотрение", type: "Оборотный", date: new Date().toISOString().slice(0, 10), notes: "" });
  const set = (k) => (e) => setF(p => ({ ...p, [k]: e.target.value }));
  return (
    <div>
      <Field label="Наименование клиента"><Input value={f.client} onChange={set("client")} placeholder="ТОО / ИП" /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="Сумма (₸)"><Input type="number" value={f.amount} onChange={set("amount")} placeholder="0" /></Field>
        <Field label="Тип финансирования"><Select value={f.type} onChange={set("type")} options={["Оборотный", "Инвестиционный", "Рефинансирование"]} /></Field>
        <Field label="Ответственный менеджер"><Select value={f.manager} onChange={set("manager")} options={managers} /></Field>
        <Field label="Стадия"><Select value={f.stage} onChange={set("stage")} options={PROJECT_STAGES} /></Field>
        <Field label="Дата поступления"><Input type="date" value={f.date} onChange={set("date")} /></Field>
      </div>
      <Field label="Примечания"><Textarea value={f.notes} onChange={set("notes")} placeholder="Краткое описание проекта..." /></Field>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <Btn variant="ghost" onClick={onClose}>Отмена</Btn>
        <Btn onClick={() => { if (f.client && f.amount) { onSave(f); onClose(); } }}>Сохранить</Btn>
      </div>
    </div>
  );
}

// ─── VIEWS ────────────────────────────────────────────────────────────────────
function Dashboard({ tasks, projects, setView }) {
  const active = projects.filter(p => p.stage !== "Выдача").length;
  const taskActive = tasks.filter(t => t.status !== "Выполнено").length;
  const overdue = tasks.filter(t => isOverdue(t.deadline, t.status)).length;
  const done = tasks.filter(t => t.status === "Выполнено").length;

  const stageCounts = PROJECT_STAGES.reduce((a, s) => { a[s] = projects.filter(p => p.stage === s).length; return a; }, {});
  const totalAmt = projects.filter(p => p.stage !== "Выдача").reduce((s, p) => s + Number(p.amount), 0);

  const managerNames = [...new Set([...projects.map(p => p.manager), ...tasks.map(t => t.manager)])].filter(Boolean);
  const byManager = managerNames.map(m => ({
    name: m,
    projects: projects.filter(p => p.manager === m && p.stage !== "Выдача").length,
    tasks: tasks.filter(t => t.manager === m && t.status !== "Выполнено").length,
    overdue: tasks.filter(t => t.manager === m && isOverdue(t.deadline, t.status)).length,
  }));

  const urgent = tasks.filter(t => t.status !== "Выполнено" && t.deadline)
    .sort((a, b) => new Date(a.deadline) - new Date(b.deadline)).slice(0, 5);

  return (
    <div>
      <div style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "'DM Serif Display', serif" }}>Дашборд</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 2 }}>Отдел кредитования МСБ · {today.toLocaleDateString("ru-KZ", { day: "numeric", month: "long", year: "numeric" })}</div>
        </div>
        <Btn onClick={() => exportToExcel(projects, tasks, "all")} variant="ghost">📥 Экспорт в Excel</Btn>
      </div>

      {/* Stat cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 24 }}>
        <StatCard icon="📁" label="Активных проектов" value={active} accent={C.info} sub={fmtM(totalAmt)} onClick={() => setView("projects")} />
        <StatCard icon="⚡" label="Задач в работе" value={taskActive} accent={C.accent} onClick={() => setView("tasks")} />
        <StatCard icon="🔴" label="Просроченных" value={overdue} accent={overdue > 0 ? C.danger : C.success} sub={overdue > 0 ? "Требуют внимания" : "Всё в срок"} />
        <StatCard icon="✅" label="Выполнено" value={done} accent={C.success} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        {/* Pipeline */}
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 14 }}>Воронка проектов</div>
          {PROJECT_STAGES.map(s => {
            const cnt = stageCounts[s];
            const pct = projects.length ? (cnt / projects.length) * 100 : 0;
            return (
              <div key={s} style={{ marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <span style={{ fontSize: 12, color: C.muted }}>{s}</span>
                  <span style={{ fontSize: 12, fontWeight: 700, color: stageColor(s) }}>{cnt}</span>
                </div>
                <div style={{ height: 6, background: C.border, borderRadius: 3 }}>
                  <div style={{ height: "100%", width: pct + "%", background: stageColor(s), borderRadius: 3, transition: "width .4s" }} />
                </div>
              </div>
            );
          })}
        </div>

        {/* Urgent tasks */}
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "18px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 14 }}>⏰ Ближайшие дедлайны</div>
          {urgent.length === 0 && <div style={{ fontSize: 12, color: C.muted }}>Нет активных задач</div>}
          {urgent.map(t => {
            const dl = daysLeft(t.deadline);
            const color = dl < 0 ? C.danger : dl === 0 ? C.warning : dl <= 2 ? C.accent : C.muted;
            return (
              <div key={t.id} style={{ display: "flex", alignItems: "flex-start", gap: 10, marginBottom: 10, paddingBottom: 10, borderBottom: `1px solid ${C.border}` }}>
                <div style={{ fontSize: 16 }}>{typeIcon(t.type)}</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: C.text, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.client}</div>
                  <div style={{ fontSize: 11, color: C.muted }}>{t.type} · {t.manager}</div>
                </div>
                <div style={{ fontSize: 11, fontWeight: 700, color, whiteSpace: "nowrap" }}>
                  {dl < 0 ? `−${Math.abs(dl)} дн` : dl === 0 ? "Сегодня" : `+${dl} дн`}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Team table */}
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "18px 20px" }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 14 }}>👥 Нагрузка по менеджерам</div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                {["Менеджер", "Проектов", "Задач", "Просрочено"].map(h =>
                  <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: C.muted, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                )}
              </tr>
            </thead>
            <tbody>
              {byManager.map(m => (
                <tr key={m.name} style={{ borderBottom: `1px solid ${C.border}22` }}>
                  <td style={{ padding: "10px 10px", color: C.text, fontWeight: 600 }}>{m.name}</td>
                  <td style={{ padding: "10px 10px" }}><Chip label={m.projects} color={C.info} /></td>
                  <td style={{ padding: "10px 10px" }}><Chip label={m.tasks} color={m.tasks > 3 ? C.warning : C.success} /></td>
                  <td style={{ padding: "10px 10px" }}><Chip label={m.overdue} color={m.overdue > 0 ? C.danger : C.success} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function Projects({ projects, managers, addProject, updateProject, deleteProject }) {
  const [modal, setModal] = useState(null); // null | "add" | {project}
  const [filter, setFilter] = useState({ manager: "Все", stage: "Все", search: "" });

  const filtered = projects.filter(p =>
    (filter.manager === "Все" || p.manager === filter.manager) &&
    (filter.stage === "Все" || p.stage === filter.stage) &&
    (!filter.search || p.client.toLowerCase().includes(filter.search.toLowerCase()))
  ).sort((a, b) => new Date(b.date) - new Date(a.date));

  const totalAmt = filtered.reduce((s, p) => s + Number(p.amount), 0);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "'DM Serif Display', serif" }}>Входящие проекты</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{filtered.length} проектов · {fmtM(totalAmt)}</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <Btn variant="ghost" onClick={() => exportToExcel(projects, [], "projects")}>📥 Excel</Btn>
          <Btn onClick={() => setModal("add")}>+ Добавить</Btn>
        </div>
      </div>

      {/* Filters */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        <input value={filter.search} onChange={e => setFilter(p => ({ ...p, search: e.target.value }))} placeholder="🔍 Поиск по клиенту..." style={{ flex: 1, minWidth: 160, background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 12px", color: C.text, fontSize: 12, outline: "none" }} />
        <select value={filter.manager} onChange={e => setFilter(p => ({ ...p, manager: e.target.value }))} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 12px", color: C.text, fontSize: 12, outline: "none" }}>
          {["Все", ...managers].map(m => <option key={m}>{m}</option>)}
        </select>
        <select value={filter.stage} onChange={e => setFilter(p => ({ ...p, stage: e.target.value }))} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 12px", color: C.text, fontSize: 12, outline: "none" }}>
          {["Все", ...PROJECT_STAGES].map(s => <option key={s}>{s}</option>)}
        </select>
      </div>

      {/* Table */}
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr style={{ background: C.surface }}>
                {["Клиент", "Сумма", "Тип", "Менеджер", "Стадия", "Дата", ""].map(h =>
                  <th key={h} style={{ textAlign: "left", padding: "10px 14px", color: C.muted, fontWeight: 600, whiteSpace: "nowrap", borderBottom: `1px solid ${C.border}` }}>{h}</th>
                )}
              </tr>
            </thead>
            <tbody>
              {filtered.map(p => (
                <tr key={p.id} style={{ borderBottom: `1px solid ${C.border}22`, cursor: "pointer" }}
                  onMouseEnter={e => e.currentTarget.style.background = C.surface + "88"}
                  onMouseLeave={e => e.currentTarget.style.background = ""}
                  onClick={() => setModal(p)}>
                  <td style={{ padding: "12px 14px", color: C.text, fontWeight: 600, maxWidth: 200, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                    {p.client}
                    {p.notes && <div style={{ fontSize: 10, color: C.muted, fontWeight: 400, marginTop: 1 }}>{p.notes}</div>}
                  </td>
                  <td style={{ padding: "12px 14px", color: C.accent, fontWeight: 700, whiteSpace: "nowrap" }}>{fmtM(p.amount)}</td>
                  <td style={{ padding: "12px 14px" }}><Chip label={p.type} color={p.type === "Инвестиционный" ? C.info : C.muted} /></td>
                  <td style={{ padding: "12px 14px", color: C.text, whiteSpace: "nowrap" }}>{p.manager}</td>
                  <td style={{ padding: "12px 14px" }}><Chip label={p.stage} color={stageColor(p.stage)} /></td>
                  <td style={{ padding: "12px 14px", color: C.muted, whiteSpace: "nowrap" }}>{new Date(p.date).toLocaleDateString("ru-KZ")}</td>
                  <td style={{ padding: "12px 14px" }}>
                    <button onClick={e => { e.stopPropagation(); setModal(p); }} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 14 }}>✏️</button>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={7} style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>Проекты не найдены</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <Modal open={!!modal} onClose={() => setModal(null)}
        title={modal === "add" ? "Новый проект" : "Редактировать проект"}>
        {modal === "add" && <ProjectForm managers={managers} onSave={async (f) => { await addProject(f); setModal(null); }} onClose={() => setModal(null)} />}
        {modal && modal !== "add" && (
          <div>
            <ProjectForm managers={managers} init={modal} onSave={async (f) => { await updateProject(modal.id, f); setModal(null); }} onClose={() => setModal(null)} />
            <div style={{ marginTop: 8, paddingTop: 12, borderTop: `1px solid ${C.border}` }}>
              <Btn variant="danger" small onClick={async () => { await deleteProject(modal.id); setModal(null); }}>🗑 Удалить проект</Btn>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}

function Tasks({ tasks, managers, addTask, updateTask, deleteTask }) {
  const [modal, setModal] = useState(null);
  const [filter, setFilter] = useState({ manager: "Все", status: "Все", type: "Все", search: "" });

  const filtered = tasks.filter(t =>
    (filter.manager === "Все" || t.manager === filter.manager) &&
    (filter.status === "Все" || t.status === filter.status) &&
    (filter.type === "Все" || t.type === filter.type) &&
    (!filter.search || t.client.toLowerCase().includes(filter.search.toLowerCase()))
  ).sort((a, b) => {
    const oa = isOverdue(a.deadline, a.status) ? -1 : 0;
    const ob = isOverdue(b.deadline, b.status) ? -1 : 0;
    if (oa !== ob) return oa - ob;
    return new Date(a.deadline || "9999") - new Date(b.deadline || "9999");
  });

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "'DM Serif Display', serif" }}>Текущие задачи</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{filtered.length} задач</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <Btn variant="ghost" onClick={() => exportToExcel([], tasks, "tasks")}>📥 Excel</Btn>
          <Btn onClick={() => setModal("add")}>+ Задача</Btn>
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        <input value={filter.search} onChange={e => setFilter(p => ({ ...p, search: e.target.value }))} placeholder="🔍 Клиент..." style={{ flex: 1, minWidth: 140, background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 12px", color: C.text, fontSize: 12, outline: "none" }} />
        {[["status", ["Все", ...TASK_STATUSES]], ["type", ["Все", ...TASK_TYPES]]].map(([k, opts]) => (
          <select key={k} value={filter[k]} onChange={e => setFilter(p => ({ ...p, [k]: e.target.value }))} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 12px", color: C.text, fontSize: 12, outline: "none" }}>
            {opts.map(o => <option key={o}>{o}</option>)}
          </select>
        ))}
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {filtered.map(t => {
          const od = isOverdue(t.deadline, t.status);
          const dl = t.deadline ? daysLeft(t.deadline) : null;
          const dlColor = od ? C.danger : dl === 0 ? C.warning : dl !== null && dl <= 2 ? C.accent : C.muted;
          return (
            <div key={t.id} onClick={() => setModal(t)}
              style={{ background: C.card, border: `1px solid ${od ? C.danger + "44" : C.border}`, borderRadius: 12, padding: "14px 16px", cursor: "pointer", display: "flex", alignItems: "flex-start", gap: 12, transition: "background .15s" }}
              onMouseEnter={e => e.currentTarget.style.background = C.surface}
              onMouseLeave={e => e.currentTarget.style.background = C.card}>
              <div style={{ fontSize: 20, flexShrink: 0, marginTop: 1 }}>{typeIcon(t.type)}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 4 }}>
                  <span style={{ fontWeight: 700, fontSize: 13, color: C.text }}>{t.client}</span>
                  <Chip label={t.type} color={C.info} />
                  <Chip label={t.status} color={statusColor(t.status)} />
                  <Chip label={t.priority} color={priorityColor(t.priority)} />
                  {od && <Chip label="Просрочено" color={C.danger} />}
                </div>
                <div style={{ fontSize: 11, color: C.muted }}>
                  {t.manager}
                  {t.notes && <span> · {t.notes}</span>}
                </div>
              </div>
              {t.deadline && (
                <div style={{ textAlign: "right", flexShrink: 0 }}>
                  <div style={{ fontSize: 11, color: dlColor, fontWeight: 700 }}>
                    {od ? `Просрочено ${Math.abs(dl)} дн` : dl === 0 ? "Сегодня" : `${dl} дн`}
                  </div>
                  <div style={{ fontSize: 10, color: C.muted }}>{new Date(t.deadline).toLocaleDateString("ru-KZ", { day: "numeric", month: "short" })}</div>
                </div>
              )}
            </div>
          );
        })}
        {filtered.length === 0 && (
          <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>Задачи не найдены</div>
        )}
      </div>

      <Modal open={!!modal} onClose={() => setModal(null)}
        title={modal === "add" ? "Новая задача" : "Редактировать задачу"}>
        {modal === "add" && <TaskForm managers={managers} onSave={addTask} onClose={() => setModal(null)} />}
        {modal && modal !== "add" && (
          <div>
            <TaskForm managers={managers} init={modal} onSave={(f) => { updateTask(modal.id, f); setModal(null); }} onClose={() => setModal(null)} />
            <div style={{ marginTop: 8, paddingTop: 12, borderTop: `1px solid ${C.border}` }}>
              <Btn variant="danger" small onClick={() => deleteTask(modal.id)}>🗑 Удалить задачу</Btn>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}

function Team({ tasks, projects, managers, addManager, removeManager }) {
  const [modal, setModal] = useState(false);
  const [newName, setNewName] = useState("");
  const [newRole, setNewRole] = useState("Менеджер");

  return (
    <div>
      <div style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "'DM Serif Display', serif" }}>Команда</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 2 }}>{managers.length} менеджеров</div>
        </div>
        <Btn onClick={() => setModal(true)}>+ Менеджер</Btn>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
        {managers.map(m => {
          const mProjects = projects.filter(p => p.manager === m && p.stage !== "Выдача");
          const mTasks = tasks.filter(t => t.manager === m && t.status !== "Выполнено");
          const mOver = mTasks.filter(t => isOverdue(t.deadline, t.status));
          const mDone = tasks.filter(t => t.manager === m && t.status === "Выполнено").length;
          return (
            <div key={m} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "18px 20px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
                <div style={{ width: 40, height: 40, borderRadius: 20, background: C.accent + "22", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 800, color: C.accent }}>
                  {m[0]}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, color: C.text, fontSize: 14 }}>{m}</div>
                  <div style={{ fontSize: 11, color: C.muted }}>Менеджер</div>
                </div>
                {mOver.length > 0 && <Chip label={`${mOver.length} просроч.`} color={C.danger} />}
                <button onClick={() => removeManager(m)} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 16, padding: 4 }} title="Удалить">🗑</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
                <div style={{ textAlign: "center", background: C.surface, borderRadius: 8, padding: "10px 6px" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: C.info }}>{mProjects.length}</div>
                  <div style={{ fontSize: 10, color: C.muted }}>Проектов</div>
                </div>
                <div style={{ textAlign: "center", background: C.surface, borderRadius: 8, padding: "10px 6px" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: C.accent }}>{mTasks.length}</div>
                  <div style={{ fontSize: 10, color: C.muted }}>Задач</div>
                </div>
                <div style={{ textAlign: "center", background: C.surface, borderRadius: 8, padding: "10px 6px" }}>
                  <div style={{ fontSize: 20, fontWeight: 800, color: C.success }}>{mDone}</div>
                  <div style={{ fontSize: 10, color: C.muted }}>Выполнено</div>
                </div>
              </div>
              {mTasks.slice(0, 3).map(t => (
                <div key={t.id} style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6, fontSize: 11 }}>
                  <span>{typeIcon(t.type)}</span>
                  <span style={{ flex: 1, color: C.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.client}</span>
                  <Chip label={t.status} color={statusColor(t.status)} />
                </div>
              ))}
              {mTasks.length > 3 && <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>+{mTasks.length - 3} ещё...</div>}
              {mTasks.length === 0 && <div style={{ fontSize: 11, color: C.muted }}>Нет активных задач</div>}
            </div>
          );
        })}
      </div>

      <Modal open={modal} onClose={() => { setModal(false); setNewName(""); }} title="Добавить менеджера">
        <Field label="Имя менеджера">
          <Input value={newName} onChange={e => setNewName(e.target.value)} placeholder="Например: Берик А." />
        </Field>
        <Field label="Должность">
          <Input value={newRole} onChange={e => setNewRole(e.target.value)} placeholder="Менеджер" />
        </Field>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <Btn variant="ghost" onClick={() => { setModal(false); setNewName(""); }}>Отмена</Btn>
          <Btn onClick={async () => { if (newName.trim()) { await addManager(newName.trim(), newRole); setModal(false); setNewName(""); } }}>Добавить</Btn>
        </div>
      </Modal>
    </div>
  );
}

// ─── APP ──────────────────────────────────────────────────────────────────────
const NAV = [
  { id: "dashboard", icon: "📊", label: "Дашборд" },
  { id: "projects", icon: "📁", label: "Проекты" },
  { id: "tasks", icon: "⚡", label: "Задачи" },
  { id: "team", icon: "👥", label: "Команда" },
];

// Detect mobile via screen width
function useIsMobile() {
  const [mobile, setMobile] = useState(window.innerWidth < 768);
  useState(() => {
    const fn = () => setMobile(window.innerWidth < 768);
    window.addEventListener("resize", fn);
    return () => window.removeEventListener("resize", fn);
  });
  return mobile;
}

export default function App() {
  const [view, setView] = useState("dashboard");
  const { projects, loading: pLoad, addProject, updateProject, deleteProject } = useProjects();
  const { tasks, loading: tLoad, addTask, updateTask, deleteTask } = useTasks();
  const { managers, addManager, removeManager } = useManagers();
  const isMobile = useIsMobile();
  const overdue = tasks.filter(t => isOverdue(t.deadline, t.status)).length;
  const loading = pLoad || tLoad;

  if (loading) return (
    <div style={{ height: "100vh", background: C.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16 }}>
      <div style={{ fontSize: 40 }}>🏦</div>
      <div style={{ fontSize: 16, color: C.accent, fontWeight: 700 }}>МСБ Кредит</div>
      <div style={{ fontSize: 12, color: C.muted }}>Загрузка данных...</div>
      <div style={{ width: 180, height: 3, background: C.border, borderRadius: 2, overflow: "hidden", marginTop: 8 }}>
        <div style={{ height: "100%", width: "60%", background: C.accent, borderRadius: 2, animation: "none" }} />
      </div>
    </div>
  );

  return (
    <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", height: "100vh", background: C.bg, color: C.text, fontFamily: "'DM Sans', system-ui, sans-serif", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@400;600;700;800&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 4px; height: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #1C3454; border-radius: 3px; }
        select option { background: #122540; }
      `}</style>

      {/* ── DESKTOP: Left Sidebar ── */}
      {!isMobile && (
        <div style={{ width: 220, background: C.surface, borderRight: `1px solid ${C.border}`, display: "flex", flexDirection: "column", flexShrink: 0 }}>
          <div style={{ padding: "20px 20px 16px", borderBottom: `1px solid ${C.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>🏦</div>
              <div>
                <div style={{ fontSize: 12, fontWeight: 800, color: C.text, lineHeight: 1 }}>МСБ Кредит</div>
                <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>Контроль отдела</div>
              </div>
            </div>
          </div>
          <nav style={{ padding: "12px 10px", flex: 1 }}>
            {NAV.map(n => {
              const active = view === n.id;
              const badge = n.id === "tasks" && overdue > 0 ? overdue : null;
              return (
                <button key={n.id} onClick={() => setView(n.id)} style={{
                  display: "flex", alignItems: "center", gap: 10, width: "100%",
                  padding: "10px 12px", borderRadius: 10, border: "none", cursor: "pointer",
                  background: active ? C.accentDim : "transparent",
                  color: active ? C.accent : C.muted,
                  fontSize: 13, fontWeight: active ? 700 : 500,
                  marginBottom: 2, transition: "all .15s", textAlign: "left",
                }}
                  onMouseEnter={e => { if (!active) e.currentTarget.style.background = C.card; }}
                  onMouseLeave={e => { if (!active) e.currentTarget.style.background = "transparent"; }}
                >
                  <span style={{ fontSize: 16 }}>{n.icon}</span>
                  {n.label}
                  {badge && <span style={{ marginLeft: "auto", background: C.danger, color: "#fff", borderRadius: 10, padding: "1px 6px", fontSize: 10, fontWeight: 800 }}>{badge}</span>}
                </button>
              );
            })}
          </nav>
          <div style={{ padding: "16px", borderTop: `1px solid ${C.border}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 30, height: 30, borderRadius: 15, background: C.accent + "33", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, color: C.accent, fontWeight: 800 }}>С</div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, color: C.text }}>Самат</div>
                <div style={{ fontSize: 10, color: C.muted }}>Руководитель</div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── MOBILE: Top Header ── */}
      {isMobile && (
        <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "12px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 28, height: 28, borderRadius: 7, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14 }}>🏦</div>
            <div style={{ fontSize: 13, fontWeight: 800, color: C.text }}>МСБ Кредит</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            {overdue > 0 && (
              <div style={{ background: C.danger, color: "#fff", borderRadius: 10, padding: "2px 8px", fontSize: 11, fontWeight: 800 }}>
                🔴 {overdue} просроч.
              </div>
            )}
            <div style={{ width: 28, height: 28, borderRadius: 14, background: C.accent + "33", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, color: C.accent, fontWeight: 800 }}>С</div>
          </div>
        </div>
      )}

      {/* ── Main Content ── */}
      <div style={{ flex: 1, overflow: "auto", padding: isMobile ? "16px 14px 80px" : "28px 28px" }}>
        {view === "dashboard" && <Dashboard tasks={tasks} projects={projects} setView={setView} />}
        {view === "projects" && <Projects projects={projects} managers={managers} addProject={addProject} updateProject={updateProject} deleteProject={deleteProject} />}
        {view === "tasks" && <Tasks tasks={tasks} managers={managers} addTask={addTask} updateTask={updateTask} deleteTask={deleteTask} />}
        {view === "team" && <Team tasks={tasks} projects={projects} managers={managers} addManager={addManager} removeManager={removeManager} />}
      </div>

      {/* ── MOBILE: Bottom Tab Bar ── */}
      {isMobile && (
        <div style={{
          position: "fixed", bottom: 0, left: 0, right: 0,
          background: C.surface, borderTop: `1px solid ${C.border}`,
          display: "flex", zIndex: 100,
          paddingBottom: "env(safe-area-inset-bottom)",
        }}>
          {NAV.map(n => {
            const active = view === n.id;
            const badge = n.id === "tasks" && overdue > 0 ? overdue : null;
            return (
              <button key={n.id} onClick={() => setView(n.id)} style={{
                flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                padding: "10px 4px 8px", border: "none", background: "transparent", cursor: "pointer",
                color: active ? C.accent : C.muted, position: "relative",
              }}>
                {active && <div style={{ position: "absolute", top: 0, left: "50%", transform: "translateX(-50%)", width: 32, height: 2, background: C.accent, borderRadius: "0 0 2px 2px" }} />}
                <span style={{ fontSize: 20, lineHeight: 1 }}>{n.icon}</span>
                <span style={{ fontSize: 10, fontWeight: active ? 700 : 500, marginTop: 3 }}>{n.label}</span>
                {badge && (
                  <span style={{ position: "absolute", top: 6, right: "calc(50% - 18px)", background: C.danger, color: "#fff", borderRadius: 8, padding: "0 4px", fontSize: 9, fontWeight: 800, minWidth: 14, textAlign: "center" }}>{badge}</span>
                )}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
