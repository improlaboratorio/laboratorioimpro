import React, { useState, useEffect, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { storage, firebaseReady } from "./storage.js";
import {
  Plus, Trash2, Check, Pencil, Users, ClipboardList, BarChart3, CalendarDays,
  ChevronLeft, ChevronRight, X, Save, AlertCircle, List, Search, Lock,
  LogOut, Sparkles, UserX, Eye, Euro, FileDown
} from "lucide-react";

/* ============ TOKENS ============ */
const T = {
  ink: "#0a0a0a",
  paper: "#faf9f7",
  paperDim: "#efece7",
  line: "#ddd8d0",
  accent: "#e31414",
  accentDim: "#fbe4e2",
  green: "#2f7a52",
  greenBg: "#e2f0e6",
  red: "#b3452c",
  redBg: "#f4e2db",
  muted: "#7a746a",
  blue: "#3462c9",
  blueBg: "#e3eafc",
  purple: "#7a3fc9",
  purpleBg: "#eee4fb",
  amber: "#b3781f",
  amberBg: "#faecd2",
};

const TEACHER_PALETTE = ["#e31414", "#0a0a0a", "#2f8f7a", "#c9457a", "#3f7ac9", "#c98f2f", "#5c8a3f", "#7a4fc9", "#c96a2f", "#2f9bc9"];

/* ============ AUTH (contraseña maestra de admin + contraseña personal por profesor) ============ */
const ADMIN_MASTER_PASSWORD = "labimpro-admin";
const PASSWORD_BASE = "labimpro";
function slugify(s) { return (s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, ""); }
function genTeacherPassword(name) { return `${PASSWORD_BASE}-${slugify(name)}-${Math.floor(1000 + Math.random() * 9000)}`; }

/* ============ GRUPOS (renombrados) ============ */
const DAYS = [
  { key: "lunes", label: "Lunes", dow: 1 },
  { key: "martes", label: "Martes", dow: 2 },
  { key: "miercoles", label: "Miércoles", dow: 3 },
  { key: "jueves", label: "Jueves", dow: 4 },
  { key: "viernes", label: "Viernes", dow: 5 },
];
const LEVELS_BY_DAY = {
  lunes: ["Iniciación 1", "Expert@s"],
  martes: ["Expert@s", "Avanzado 1"],
  miercoles: ["Intermedio 1", "Veteranos"],
  jueves: ["Iniciación 2", "Iniciación 3"],
  viernes: ["Intermedio 2", "Avanzado 2"],
};
const GROUPS = DAYS.flatMap((d) => LEVELS_BY_DAY[d.key].map((level) => ({
  id: `${d.key}-${level.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
  day: d.key,
  dayLabel: d.label,
  dow: d.dow,
  level,
  label: `${d.label} · ${level}`,
})));

/* Clases sueltas / esporádicas: además de los 10 grupos fijos de arriba, el admin
   puede crear clases puntuales (individuales o grupales) cualquier día de la semana,
   incluidos sábado y domingo. Se guardan en roster.clasesSueltas y se buscan con
   findGroup() junto con los grupos fijos, así el resto de la app (ficha de clase,
   listados...) no necesita saber si un grupo es "fijo" o "suelto".            */
function findGroup(roster, groupId) {
  return GROUPS.find((g) => g.id === groupId) || (roster.clasesSueltas || []).find((g) => g.id === groupId);
}
// Abreviatura corta del nivel para mostrar junto al profesor en las casillas del calendario
// (p.ej. "Iniciación 1" -> "INIC 1", "Avanzado 2" -> "AVAN 2", "Expert@s" -> "EXP").
function levelAbbrev(level) {
  if (!level) return "";
  const norm = level.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const m = norm.match(/^(Iniciacion|Avanzado|Intermedio)\s*(\d+)?/i);
  if (m) {
    const key = m[1].toLowerCase();
    const abbr = key === "iniciacion" ? "INIC" : key === "avanzado" ? "AVAN" : "INTER";
    return m[2] ? `${abbr} ${m[2]}` : abbr;
  }
  if (/^expert/i.test(norm)) return "EXP";
  if (/^veteranos/i.test(norm)) return "VET";
  return level.length <= 8 ? level.toUpperCase() : level.slice(0, 6).toUpperCase();
}
const ALL_LEVELS = [...new Set(GROUPS.map((g) => g.level))];

const MONTH_NAMES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
const DOW_NAMES = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];

function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }
function pad2(n) { return String(n).padStart(2, "0"); }
function fileToResizedDataURL(file, maxSize = 200, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("No se pudo leer el archivo"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("No se pudo leer la imagen"));
      img.onload = () => {
        const side = Math.min(img.width, img.height);
        const sx = (img.width - side) / 2, sy = (img.height - side) / 2;
        const canvas = document.createElement("canvas");
        canvas.width = maxSize; canvas.height = maxSize;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, sx, sy, side, side, 0, 0, maxSize, maxSize);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}
function iso(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function parseISO(s) { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); }
function fmtDateShort(s) { const d = parseISO(s); return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}`; }
function fmtDateLong(s) { const d = parseISO(s); return d.toLocaleDateString("es-ES", { weekday: "long", day: "2-digit", month: "long", year: "numeric" }); }
function yyyymm(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`; }
function monthLabel(ym) { const [y, m] = ym.split("-").map(Number); return `${MONTH_NAMES[m - 1]} ${y}`; }

function lastWeekdayOnOrBefore(from, dow) {
  const d = new Date(from);
  const diff = (d.getDay() - dow + 7) % 7;
  d.setDate(d.getDate() - diff);
  return d;
}
function weekdayColumnsForRange(dow, monthsBack) {
  const now = new Date();
  const rangeStart = new Date(now.getFullYear(), now.getMonth() - monthsBack, 1);
  const rangeEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const cols = [];
  let d = new Date(rangeStart);
  const offset = (dow - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + offset);
  while (d <= rangeEnd) {
    cols.push(iso(d));
    d = new Date(d);
    d.setDate(d.getDate() + 7);
  }
  return cols;
}
const LOGO_SRC = "data:image/png;base64,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";


/* ============ UI PRIMITIVES ============ */
function Card({ children, style }) {
  return <div className="app-card" style={{ background: "#fff", border: `1px solid ${T.line}`, borderTop: `3px solid ${T.accentDim}`, borderRadius: 16, padding: 20, marginBottom: 16, boxShadow: "0 1px 2px rgba(10,10,10,0.04), 0 6px 16px -8px rgba(10,10,10,0.08)", ...style }}>{children}</div>;
}
function Label({ children }) {
  return <div style={{ fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase", color: T.muted, fontWeight: 700, marginBottom: 7 }}>{children}</div>;
}
function inputStyle(extra) {
  return { width: "100%", padding: "10px 12px", border: `1px solid ${T.line}`, borderRadius: 10, fontSize: 14, background: T.paper, color: T.ink, transition: "border-color .15s", ...extra };
}
function btn(variant, extra) {
  const base = { display: "inline-flex", alignItems: "center", gap: 6, padding: "9px 15px", borderRadius: 10, fontSize: 13.5, fontWeight: 700, border: "1px solid transparent" };
  if (variant === "primary") return { ...base, background: `linear-gradient(135deg, ${T.accent}, #c9457a)`, color: "#fff", boxShadow: "0 4px 10px -4px rgba(227,20,20,0.55)", ...extra };
  if (variant === "ghost") return { ...base, background: "#fff", color: T.ink, border: `1px solid ${T.line}`, ...extra };
  if (variant === "danger") return { ...base, background: "transparent", color: T.red, border: `1px solid ${T.redBg}`, ...extra };
  return { ...base, ...extra };
}
function Avatar({ name, photo, size = 32, color }) {
  const initial = (name || "?").trim()[0]?.toUpperCase() || "?";
  if (photo) {
    return <img src={photo} alt={name} style={{ width: size, height: size, borderRadius: "50%", objectFit: "cover", flexShrink: 0, border: `1px solid ${T.line}` }} />;
  }
  return (
    <div style={{
      width: size, height: size, borderRadius: "50%", flexShrink: 0, background: color || T.paperDim,
      color: color ? "#fff" : T.muted, display: "flex", alignItems: "center", justifyContent: "center",
      fontWeight: 800, fontSize: Math.round(size * 0.4),
    }}>{initial}</div>
  );
}
function Badge({ children, tone = "neutral" }) {
  const tones = {
    neutral: { background: T.paperDim, color: T.muted },
    accent: { background: T.accentDim, color: T.accent },
    green: { background: T.greenBg, color: T.green },
    blue: { background: T.blueBg, color: T.blue },
    purple: { background: T.purpleBg, color: T.purple },
    amber: { background: T.amberBg, color: T.amber },
  };
  const c = tones[tone] || tones.neutral;
  return <span style={{ display: "inline-block", padding: "3px 10px", borderRadius: 999, fontSize: 10.5, fontWeight: 800, letterSpacing: "0.03em", textTransform: "uppercase", ...c }}>{children}</span>;
}
function attendanceBoxStyle(present, size) {
  return {
    width: size, height: size, borderRadius: Math.round(size / 3), flexShrink: 0,
    border: `1px solid ${present ? T.green : T.line}`,
    background: present ? T.green : "#fff",
    display: "flex", alignItems: "center", justifyContent: "center",
  };
}
function SubTabs({ tabs, active, setActive }) {
  return (
    <div style={{ display: "flex", gap: 4, flexWrap: "wrap", background: T.paperDim, borderRadius: 10, padding: 4, marginBottom: 16 }}>
      {tabs.map((t) => (
        <button key={t.id} onClick={() => setActive(t.id)} style={{
          padding: "7px 13px", border: "none", background: active === t.id ? "#fff" : "transparent", fontSize: 12.5, borderRadius: 8,
          fontWeight: active === t.id ? 800 : 600, color: active === t.id ? T.ink : T.muted,
          boxShadow: active === t.id ? "0 1px 3px rgba(10,10,10,0.12)" : "none",
        }}>{t.label}</button>
      ))}
    </div>
  );
}

function PasswordField({ value, onSave }) {
  const [val, setVal] = useState(value || "");
  const [show, setShow] = useState(false);
  useEffect(() => setVal(value || ""), [value]);
  return (
    <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
      <div style={{ position: "relative", flex: 1 }}>
        <input
          type={show ? "text" : "password"}
          style={inputStyle({ paddingRight: 34 })}
          value={val}
          onChange={(e) => setVal(e.target.value)}
          onBlur={() => { if (val.trim()) onSave(val.trim()); }}
          onKeyDown={(e) => { if (e.key === "Enter" && val.trim()) onSave(val.trim()); }}
        />
        <button type="button" onClick={() => setShow(!show)} style={{ position: "absolute", right: 6, top: 0, bottom: 0, border: "none", background: "transparent", color: T.muted, display: "flex", alignItems: "center" }}>
          <Eye size={15} />
        </button>
      </div>
      <button style={btn("ghost", { padding: "9px 12px" })} onClick={() => val.trim() && onSave(val.trim())}>Guardar</button>
    </div>
  );
}

/* ============ DATOS POR DEFECTO ============ */
const EMPTY_ROSTER = { teachers: [], students: {}, dinamicasCatalog: [], escenasCatalog: [], tiposClase: [] };
const EMPTY_PAYMENTS = { students: {}, teachers: {} }; // students: {studentId:{yyyymm:{paid,amount,paidTo}}}, teachers: {name:{yyyymm:{paid}}}
const EMPTY_PENDING = {}; // { groupId: [dinamicaId,...] }

/* Catálogo importado de ejemplo (5 tipos de clase, del documento de referencia) — solo se usa
   si todavía no hay ningún tipo de clase guardado, como punto de partida editable. */
const SEED_DATA = {"dinamicasCatalog": [{"id": "din-1", "text": "RUEDA DE NOMBRES CON GESTO EN CADA SÍLABA. LUEGO TODOS LO REPETIMOS"}, {"id": "din-2", "text": "EN CÍRCULO ENVIAR Y RECIBIR PALMADA"}, {"id": "din-3", "text": "EN CIRCULO ENVIAR NOMBRE"}, {"id": "din-4", "text": "ANDAR A VELOCIDADES (INCLUIMOS -2)"}, {"id": "din-5", "text": "PARAMOS Y ANDAMOS JUSTO CON LA INDICACION CONTRARIA"}, {"id": "din-6", "text": "1-SALTO Y PALMADA, 2- AGACHAMOS, 3- GIRO, 4- CAMBIO DIRECCIÓN, 5- SALUDO"}, {"id": "din-7", "text": "BOLA DE ENERGÍA (HIA, JONDOM, AYYY, LA TIRO)"}, {"id": "din-8", "text": "EN CÍRCULO LANZAR CONTRARIOS BÁSICO (SI NO IZQUIERDA DERECHA)"}, {"id": "din-9", "text": "VOLCÁN DE PALABRAS CON LO QUE SUGIERE"}, {"id": "din-10", "text": "SI TU ERES TAL YO SOY TAL Y VOLVER"}, {"id": "din-11", "text": "FOTOGRAFÍA HUMANA – SOY… TÍTULO Y ME QUEDO CON…"}, {"id": "din-12", "text": "CONTAR DEL 1 AL 10 SIN COINCIDIR Y BAJAR"}, {"id": "din-13", "text": "1000 MANERAS DE SALUDAR – 1 SOLO MIRADA – 2 GESTO – 3 HACE MUCHO TIEMPO Y ESTAIS MUY CONTENTOS DE VEROS – POLÍTICOS – CON VERGÜENZA (CAMBIADA POR LA SIGUIENTE)"}, {"id": "din-14", "text": "Saludo secreto, objeto secreto, canción favorita y amigo de la infancia."}, {"id": "din-15", "text": "EN CIRCULO ENVIAR NOMBRE (varios niveles, de simple a normal)"}, {"id": "din-16", "text": "BOLA DE ENERGÍA"}, {"id": "din-17", "text": "PASILLO"}, {"id": "din-18", "text": "3 COSAS 3"}, {"id": "din-19", "text": "ESPEJO (POR PAREJAS)"}, {"id": "din-20", "text": "SÍ Y ADEMÁS (POR PAREJAS – PRIMERO CON OBJETOS REALES, DESPUES MIMADOS)"}, {"id": "din-21", "text": "ANDAR DECIR FREEZE, CERRAR OJOS Y PREGUNTARLES COSAS DE ELLOS FÍSICAS, EN PLAN, CUANTOS SOMOS?, CUANTOS LLEVAN GAFAS…"}, {"id": "din-22", "text": "SOLO UNO ANDANDO Y SE PARA DELANTE DE OTRO Y ESE OTRO EMPIEZA A ANDAR (OPCIÓN DE IR CANTANDO ALGO)"}, {"id": "din-23", "text": "1 ANDANDO, 2 ANDANDO, 1 AGACHADO…"}, {"id": "din-24", "text": "VOLCÁN DE PALABRAS ÚLTIMA SÍLABA"}, {"id": "din-25", "text": "EN CIRCULO ENVIAR Y RECIBIR PALMADA"}, {"id": "din-26", "text": "ENVIAR NOMBRE DEL QUE VEMOS"}, {"id": "din-27", "text": "ENVIAR NOMBRE DE OTRO"}, {"id": "din-28", "text": "JUNTARSE POR NÚMERO"}, {"id": "din-29", "text": "CABAÑA, CASA. CHALET, MANSIÓN, NÚMEROS, INQUILINOS. SOY NORMA DUVAL"}, {"id": "din-30", "text": "GIFS"}, {"id": "din-31", "text": "EL 7 SIMPLE (CON VUELTA GRITANDO Y ÚLTIMA ELIMINANDO)"}, {"id": "din-32", "text": "VOLCÁN DE PALABRAS"}, {"id": "din-33", "text": "PINTAR UN CUADRO ENTRE TODOS"}, {"id": "din-34", "text": "7 SIMPLE"}, {"id": "din-35", "text": "SERPIENTE COPIA AL DE DELANTE"}, {"id": "din-36", "text": "VOLCÁN ÚLTIMA SÍLABA"}, {"id": "din-37", "text": "LANZAR NOMBRE DEL QUE VEMOS (ASI TE LOS APRENDES)"}, {"id": "din-38", "text": "LANZAR NOMBRE DE OTRO"}, {"id": "din-39", "text": "LANZAR CONTRARIOS"}, {"id": "din-40", "text": "1,2,3 PALABRA…"}, {"id": "din-41", "text": "SI TU ERES TAL, YO SOY TAL Y VOLVER"}, {"id": "din-42", "text": "ME DAS PERMISO? SI TE DOY PERMISO. SOLO CON GESTO."}, {"id": "din-43", "text": "USA-USA, ÑAKI-ÑAKI, TOKI- TOKI"}], "tiposClase": [{"numero": 1, "nivel": "Iniciación 1", "nombre": "HOLA IMPRO", "dinamicaIds": ["din-1", "din-2", "din-3", "din-4", "din-5", "din-6", "din-7", "din-8", "din-9", "din-10", "din-11"], "escenas": "¿QUÉ ESTÁS HACIENDO?"}, {"numero": 2, "nivel": "Iniciación 1", "nombre": "SÍ Y ADEMÁS", "dinamicaIds": ["din-12", "din-13", "din-14", "din-6", "din-15", "din-16", "din-9", "din-17", "din-18", "din-19", "din-20"], "escenas": "QUÉ ESTÁS HACIENDO?\nFREEZE"}, {"numero": 3, "nivel": "Iniciación 1", "nombre": "ARQUEROS", "dinamicaIds": ["din-21", "din-22", "din-23", "din-9", "din-24", "din-25", "din-26", "din-27", "din-28", "din-29", "din-30", "din-31", "din-17"], "escenas": "¿QUÉ ESTÁS HACIENDO?\nFREEZE\nPREVIA ARQUERO SITIO DADO\nARQUERO\nARQUERO DOBLE"}, {"numero": 4, "nivel": "Iniciación 1", "nombre": "NO HAY ERROR", "dinamicaIds": ["din-32", "din-33", "din-34", "din-25", "din-26", "din-27", "din-30"], "escenas": "QUE ESTAS HACIENDO?\nARQUERO"}, {"numero": 5, "nivel": "Iniciación 1", "nombre": "ZOOM EN CONVERSACIONES", "dinamicaIds": ["din-35", "din-36", "din-37", "din-38", "din-39", "din-40", "din-41", "din-42", "din-43"], "escenas": "¿QUÉ ESTÁS HACIENDO? (1 RONDA)\nVEMOS ESCENARIO EN 3,2,1,,,,YA:\nRESTAURANTE, GIMNASIO, PISCINA, IKEA, PARQUE\nPRIMERO SE PINTA EL LUGAR Y CUANDO LE DE AL TIMBRE SE PUEDE EMPEZAR A HABLAR Y A ROBARSE EL FOCO EN ESCENA"}]};


/* ============ LOGIN ============ */
function LoginGate({ roster, onLogin }) {
  const [pwd, setPwd] = useState("");
  const [showPwd, setShowPwd] = useState(false);
  const [err, setErr] = useState("");

  function submit() {
    const p = pwd.trim().toLowerCase();
    if (!p) return;
    if (p === ADMIN_MASTER_PASSWORD.toLowerCase()) { onLogin({ role: "admin", profesorName: null }); return; }
    const teacher = roster.teachers.find((t) => (t.password || "").toLowerCase() === p);
    if (teacher) { onLogin({ role: teacher.isAdmin ? "admin" : "profesor", profesorName: teacher.name }); return; }
    setErr("Contraseña incorrecta.");
  }

  return (
    <div className="app-shell" style={{ minHeight: "100vh", background: T.paper, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "system-ui, sans-serif" }}>
      <div style={{ background: "#fff", border: `1px solid ${T.line}`, borderRadius: 14, padding: 28, width: "100%", maxWidth: 380 }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 18 }}>
          <img src={LOGO_SRC} alt="Laboratorio Impro" style={{ height: 52, width: "auto" }} />
        </div>
        <div style={{ textAlign: "center", fontSize: 13, color: T.muted, marginBottom: 14 }}>Introduce tu contraseña</div>
        <div style={{ position: "relative", marginBottom: 10 }}>
          <input autoFocus type={showPwd ? "text" : "password"} style={inputStyle({ paddingRight: 38 })} value={pwd} onChange={(e) => setPwd(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
          <button type="button" onClick={() => setShowPwd(!showPwd)} style={{ position: "absolute", right: 8, top: 0, bottom: 0, border: "none", background: "transparent", color: T.muted, display: "flex", alignItems: "center" }}>
            <Eye size={16} />
          </button>
        </div>
        {err && <div style={{ color: T.red, fontSize: 12.5, marginBottom: 10 }}>{err}</div>}
        <button style={btn("primary", { width: "100%", justifyContent: "center", padding: 12 })} onClick={submit}><Lock size={15} /> Entrar</button>
      </div>
    </div>
  );
}

/* ============ APP PRINCIPAL ============ */
export default function AsistenciaApp() {
  const [roster, setRoster] = useState(EMPTY_ROSTER);
  const [sessions, setSessions] = useState([]);
  const [payments, setPayments] = useState(EMPTY_PAYMENTS);
  const [pending, setPending] = useState(EMPTY_PENDING);
  const [loaded, setLoaded] = useState(false);
  const [error, setError] = useState("");
  const [auth, setAuth] = useState(null); // {role, profesorName?}
  const [tab, setTab] = useState("calendario");
  const [focusGroupId, setFocusGroupId] = useState(null);
  function goToAlumnosGroup(groupId) { setFocusGroupId(groupId); setTab("alumnos"); }
  const [confirmDialog, setConfirmDialog] = useState(null); // {message, onConfirm}
  function askConfirm(message, onConfirm) { setConfirmDialog({ message, onConfirm }); }

  useEffect(() => {
    (async () => {
      try {
        const r = await storage.get("asis_roster");
        const loaded = r && r.value ? JSON.parse(r.value) : {};
        const merged = { ...EMPTY_ROSTER, ...loaded };
        // Si el catálogo de tipos de clase o de dinámicas todavía está vacío (primera vez, o un
        // roster guardado antes de que existiera esta función), sembramos el catálogo de ejemplo.
        if (!merged.tiposClase || merged.tiposClase.length === 0) merged.tiposClase = SEED_DATA.tiposClase;
        if (!merged.dinamicasCatalog || merged.dinamicasCatalog.length === 0) merged.dinamicasCatalog = SEED_DATA.dinamicasCatalog;
        // Las escenas antes eran texto libre por tipo de clase; la primera vez, construimos
        // el catálogo de escenas (fijas, con check) a partir de esas líneas ya escritas.
        if (!merged.escenasCatalog || merged.escenasCatalog.length === 0) {
          const seen = new Map();
          (merged.tiposClase || []).forEach((t) => {
            (t.escenas || "").split("\n").map((l) => l.trim()).filter(Boolean).forEach((line) => {
              const key = line.toLowerCase();
              if (!seen.has(key)) seen.set(key, { id: `esc-${seen.size + 1}`, text: line });
            });
          });
          merged.escenasCatalog = [...seen.values()];
        }
        setRoster(merged);
        if (!loaded.tiposClase || !loaded.dinamicasCatalog || !loaded.escenasCatalog) {
          // guardamos ya el catálogo sembrado, para no depender de que ocurra otro guardado más tarde
          storage.set("asis_roster", JSON.stringify(merged)).catch(() => {});
        }
      } catch (e) {}
      try {
        const s = await storage.get("asis_sessions");
        if (s && s.value) setSessions(JSON.parse(s.value));
      } catch (e) {}
      try {
        const p = await storage.get("asis_payments");
        if (p && p.value) setPayments({ ...EMPTY_PAYMENTS, ...JSON.parse(p.value) });
      } catch (e) {}
      try {
        const pd = await storage.get("asis_pending");
        if (pd && pd.value) setPending(JSON.parse(pd.value));
      } catch (e) {}
      try {
        const saved = localStorage.getItem("asis_auth_v1");
        if (saved) setAuth(JSON.parse(saved));
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  const persistRoster = useCallback(async (next) => {
    setRoster(next);
    try {
      const res = await storage.set("asis_roster", JSON.stringify(next));
      setError(res ? "" : "No se ha podido guardar. Inténtalo de nuevo.");
    } catch (e) { setError("No se ha podido guardar. Comprueba la conexión."); }
  }, []);
  const persistSessions = useCallback(async (next) => {
    setSessions(next);
    try {
      const res = await storage.set("asis_sessions", JSON.stringify(next));
      setError(res ? "" : "No se ha podido guardar. Inténtalo de nuevo.");
    } catch (e) { setError("No se ha podido guardar. Comprueba la conexión."); }
  }, []);
  const persistPayments = useCallback(async (next) => {
    setPayments(next);
    try {
      const res = await storage.set("asis_payments", JSON.stringify(next));
      setError(res ? "" : "No se ha podido guardar. Inténtalo de nuevo.");
    } catch (e) { setError("No se ha podido guardar. Comprueba la conexión."); }
  }, []);
  const persistPending = useCallback(async (next) => {
    setPending(next);
    try {
      const res = await storage.set("asis_pending", JSON.stringify(next));
      setError(res ? "" : "No se ha podido guardar. Inténtalo de nuevo.");
    } catch (e) { setError("No se ha podido guardar. Comprueba la conexión."); }
  }, []);

  function addTeacher(name, color, password) {
    const n = (name || "").trim();
    if (!n || roster.teachers.some((t) => t.name === n)) return;
    const c = color || TEACHER_PALETTE[roster.teachers.length % TEACHER_PALETTE.length];
    const pwd = (password || "").trim() || genTeacherPassword(n);
    persistRoster({ ...roster, teachers: [...roster.teachers, { name: n, color: c, isAdmin: false, password: pwd }] });
  }
  function updateTeacher(name, patch) {
    persistRoster({ ...roster, teachers: roster.teachers.map((t) => (t.name === name ? { ...t, ...patch } : t)) });
  }
  function regenerateTeacherPassword(name) {
    updateTeacher(name, { password: genTeacherPassword(name) });
  }
  function removeTeacher(name) {
    persistRoster({ ...roster, teachers: roster.teachers.filter((t) => t.name !== name) });
  }
  function addDinamica(text) {
    const t = (text || "").trim();
    if (!t) return null;
    const item = { id: uid(), text: t };
    persistRoster({ ...roster, dinamicasCatalog: [...roster.dinamicasCatalog, item] });
    return item;
  }
  function removeDinamica(id) {
    persistRoster({ ...roster, dinamicasCatalog: roster.dinamicasCatalog.filter((d) => d.id !== id) });
  }
  function addEscena(text) {
    const t = (text || "").trim();
    if (!t) return null;
    const item = { id: uid(), text: t };
    persistRoster({ ...roster, escenasCatalog: [...(roster.escenasCatalog || []), item] });
    return item;
  }
  function removeEscena(id) {
    persistRoster({ ...roster, escenasCatalog: (roster.escenasCatalog || []).filter((d) => d.id !== id) });
  }
  function saveTipoClase(tipo) {
    const exists = roster.tiposClase.some((t) => t.numero === tipo.numero);
    const next = exists ? roster.tiposClase.map((t) => (t.numero === tipo.numero ? tipo : t)) : [...roster.tiposClase, tipo];
    persistRoster({ ...roster, tiposClase: next.sort((a, b) => a.numero - b.numero) });
  }
  function removeTipoClase(numero) {
    persistRoster({ ...roster, tiposClase: roster.tiposClase.filter((t) => t.numero !== numero) });
  }
  function resetSessions() {
    persistSessions([]);
    persistPending({});
  }

  function login(a) {
    setAuth(a);
    localStorage.setItem("asis_auth_v1", JSON.stringify(a));
    setTab("calendario");
  }
  function logout() {
    setAuth(null);
    localStorage.removeItem("asis_auth_v1");
  }

  if (!loaded) {
    return <div style={{ minHeight: "100vh", background: T.paper, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "system-ui", color: T.muted }}>Cargando…</div>;
  }
  if (!auth) {
    return <LoginGate roster={roster} onLogin={login} />;
  }

  const isAdmin = auth.role === "admin";

  const tabs = isAdmin
    ? [
        { id: "calendario", label: "Calendario", icon: CalendarDays },
        { id: "alumnos", label: "Alumnos", icon: Users },
        { id: "pagos", label: "Pagos", icon: Euro },
        { id: "clases", label: "Clases", icon: List },
        { id: "profesores", label: "Profesores", icon: BarChart3 },
        { id: "dinamicas", label: "Dinámicas", icon: Sparkles },
      ]
    : [{ id: "calendario", label: "Calendario", icon: CalendarDays }];

  const ctx = { roster, sessions, payments, pending, persistRoster, persistSessions, persistPayments, persistPending, addTeacher, updateTeacher, regenerateTeacherPassword, removeTeacher, addDinamica, removeDinamica, addEscena, removeEscena, saveTipoClase, removeTipoClase, resetSessions, auth, isAdmin, askConfirm, focusGroupId, setFocusGroupId, goToAlumnosGroup };

  return (
    <div className="app-shell" style={{ minHeight: "100vh", background: `linear-gradient(180deg, ${T.accentDim} 0%, ${T.paper} 220px)`, fontFamily: "system-ui, sans-serif", color: T.ink }}>
      <style>{`
        * { box-sizing: border-box; }
        button { font-family: inherit; cursor: pointer; }
        input, select, textarea { font-family: inherit; }
        ::selection { background: ${T.accentDim}; }
        table { border-collapse: separate; border-spacing: 0; }
        html { -webkit-text-size-adjust: 100%; }
        body { overflow-x: hidden; }
        .tab-scroll { display: flex; gap: 2px; overflow-x: auto; -webkit-overflow-scrolling: touch; scrollbar-width: none; }
        .tab-scroll::-webkit-scrollbar { display: none; }
        .tab-scroll button { flex-shrink: 0; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
        @media (max-width: 520px) {
          .app-card { padding: 13px; }
          .two-col { grid-template-columns: 1fr; }
          .tab-scroll button { padding: 8px 10px !important; font-size: 12.5px !important; }
          .tab-scroll button span.tab-label { display: none; }
          h1 { font-size: 20px !important; }
          .cal-cell { min-height: 84px !important; }
          .cal-day-num { font-size: 14px !important; padding: 5px 7px !important; }
          .cal-init { font-size: 10px !important; }
          .cal-dow { font-size: 12.5px !important; }
          .cal-legend { font-size: 13px !important; }
          .modal-sheet { max-height: 85vh !important; }
        }
      `}</style>

      <header style={{ background: "#fff", boxShadow: "0 1px 2px rgba(10,10,10,0.04), 0 8px 20px -12px rgba(10,10,10,0.15)", padding: "14px 12px 0" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 10.5, letterSpacing: "0.16em", textTransform: "uppercase", color: T.accent, fontWeight: 700, marginBottom: 4, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
              {isAdmin ? (auth.profesorName ? `Administrador/a · ${auth.profesorName}` : "Administrador/a") : `Profesor/a · ${auth.profesorName}`}
            </div>
            <h1 style={{ margin: "0 0 12px", fontSize: 24, fontWeight: 800 }}>{tabs.find((t) => t.id === tab)?.label || "Asistencia"}</h1>
          </div>
          <img src={LOGO_SRC} alt="Laboratorio Impro" style={{ height: 38, width: "auto", flexShrink: 0, marginTop: 2 }} />
        </div>
        <div style={{ maxWidth: 1100, margin: "0 auto" }}>
          {!firebaseReady && (
            <div style={{ display: "flex", alignItems: "center", gap: 6, color: T.accent, fontSize: 12, marginBottom: 10 }}>
              <AlertCircle size={13} /> Sin Firebase configurado: los datos solo se guardan en este navegador, no se comparten con los demás todavía. Mira <code>src/storage.js</code>.
            </div>
          )}
          {error && <div style={{ display: "flex", alignItems: "center", gap: 6, color: T.red, fontSize: 12.5, marginBottom: 10 }}><AlertCircle size={14} /> {error}</div>}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
            <div className="tab-scroll">
              {tabs.map((it) => {
                const Icon = it.icon;
                const active = tab === it.id;
                return (
                  <button key={it.id} onClick={() => setTab(it.id)} className="app-tab" style={{
                    display: "flex", alignItems: "center", gap: 6, padding: "9px 14px",
                    border: "none", background: active ? `linear-gradient(135deg, ${T.ink}, #3a2f4a)` : "transparent", fontSize: 13.5, fontWeight: active ? 700 : 500,
                    color: active ? "#fff" : T.muted, borderRadius: "10px 10px 0 0", marginBottom: -1,
                    whiteSpace: "nowrap",
                  }}><Icon size={15} /> <span className="tab-label">{it.label}</span></button>
                );
              })}
            </div>
            <button onClick={logout} style={{ display: "flex", alignItems: "center", gap: 5, border: "none", background: "transparent", color: T.muted, fontSize: 12, padding: "8px 6px", marginBottom: 4, flexShrink: 0 }}>
              <LogOut size={13} />
            </button>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1100, margin: "0 auto", padding: "18px 16px 80px" }}>
        {tab === "calendario" && <CalendarioTab ctx={ctx} />}
        {isAdmin && tab === "alumnos" && <AlumnosTab ctx={ctx} />}
        {isAdmin && tab === "pagos" && <PagosTab ctx={ctx} />}
        {isAdmin && tab === "clases" && <ClasesTab ctx={ctx} />}
        {isAdmin && tab === "profesores" && <ProfesoresTab ctx={ctx} />}
        {isAdmin && tab === "dinamicas" && <DinamicasTab ctx={ctx} />}
      </main>

      {confirmDialog && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 20 }} onClick={() => setConfirmDialog(null)}>
          <div style={{ background: "#fff", borderRadius: 14, padding: 22, width: "100%", maxWidth: 360 }} onClick={(e) => e.stopPropagation()}>
            <div style={{ fontSize: 14.5, marginBottom: 18, lineHeight: 1.5 }}>{confirmDialog.message}</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={btn("ghost", { flex: 1, justifyContent: "center" })} onClick={() => setConfirmDialog(null)}>Cancelar</button>
              <button style={btn("danger", { flex: 1, justifyContent: "center", border: `1px solid ${T.red}` })} onClick={() => { confirmDialog.onConfirm(); setConfirmDialog(null); }}>Confirmar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============ PENDIENTES DE DINÁMICAS (por grupo) ============ */
// Estado explícito de la clase: si la sesión ya tiene un campo "estado" guardado se respeta tal
// cual (para poder marcar una clase futura como impartida, o una de hoy como programada, a mano);
// si no lo tiene (sesiones antiguas), se deduce de si ya se registró alguna asistencia.
function isGiven(session) {
  if (!session) return false;
  if (session.estado) return session.estado === "impartida";
  return Object.keys(session.attendance || {}).length > 0;
}

function getPending(pending, groupId) { return pending[groupId] || []; }
function applyDynamicsToPending(pending, groupId, dinamicasList) {
  const cur = new Set(getPending(pending, groupId));
  (dinamicasList || []).forEach((d) => {
    if (d.done) cur.delete(d.id);
    else cur.add(d.id);
  });
  return { ...pending, [groupId]: [...cur] };
}
function removeFromPending(pending, groupId, dinamicaId) {
  const cur = getPending(pending, groupId).filter((id) => id !== dinamicaId);
  return { ...pending, [groupId]: cur };
}

/* ============ CALENDARIO ============ */
function CalendarioTab({ ctx }) {
  const { roster, sessions, isAdmin, auth, persistRoster, persistSessions, addDinamica, addEscena, goToAlumnosGroup } = ctx;
  const [cursor, setCursor] = useState(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  const [pickerDate, setPickerDate] = useState(null);
  const [openSession, setOpenSession] = useState(null); // {date, groupId}
  const [toast, setToast] = useState("");
  const [creatingSporadic, setCreatingSporadic] = useState(false);
  const [sporadicDraft, setSporadicDraft] = useState(() => ({ nombre: "", profesor: "", estado: "programada", dinamicas: [], escenas: [], titulos: "", conceptosFinales: "", observaciones: "" }));
  const [newDinSporadic, setNewDinSporadic] = useState("");
  const [newEscSporadic, setNewEscSporadic] = useState("");

  const byDate = useMemo(() => {
    const m = {};
    sessions.forEach((s) => { (m[s.date] = m[s.date] || []).push(s); });
    return m;
  }, [sessions]);

  const teacherColor = (name) => roster.teachers.find((t) => t.name === name)?.color || T.muted;

  const cyear = cursor.getFullYear(), cmonth = cursor.getMonth();
  const firstDay = new Date(cyear, cmonth, 1);
  const startOffset = (firstDay.getDay() + 6) % 7;
  const daysInMonth = new Date(cyear, cmonth + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  // Grupos que le corresponde ver a este usuario ese día de la semana (todos si es admin,
  // solo los que ya tienen una clase asignada a su nombre si es profesor/a).
  function dayGroupsFor(dateStr) {
    const dow = parseISO(dateStr).getDay();
    const all = GROUPS.filter((g) => g.dow === dow);
    if (isAdmin) return all;
    const mine = (byDate[dateStr] || []).filter((s) => s.profesor === auth.profesorName).map((s) => s.groupId);
    return all.filter((g) => mine.includes(g.id));
  }

  function handleDayClick(dateStr) {
    const dayGroups = dayGroupsFor(dateStr);
    const sporadicToday = (byDate[dateStr] || []).filter((s) => !GROUPS.some((g) => g.id === s.groupId));
    if (dayGroups.length === 0 && sporadicToday.length === 0) {
      if (!isAdmin) { setToast("No tienes ninguna clase asignada este día."); setTimeout(() => setToast(""), 2200); return; }
      setPickerDate(dateStr);
      return;
    }
    if (dayGroups.length === 1 && sporadicToday.length === 0) { setOpenSession({ date: dateStr, groupId: dayGroups[0].id }); return; }
    setPickerDate(dateStr);
  }

  function resetSporadicDraft() {
    setSporadicDraft({ nombre: "", profesor: "", estado: "programada", dinamicas: [], escenas: [], titulos: "", conceptosFinales: "", observaciones: "" });
    setNewDinSporadic("");
    setNewEscSporadic("");
  }
  function addDinToSporadicDraft() {
    const text = newDinSporadic.trim();
    if (!text) return;
    const existingCat = roster.dinamicasCatalog.find((c) => c.text.toLowerCase() === text.toLowerCase());
    const item = existingCat || addDinamica(text);
    if (!item || sporadicDraft.dinamicas.some((d) => d.id === item.id)) { setNewDinSporadic(""); return; }
    setSporadicDraft((d) => ({ ...d, dinamicas: [...d.dinamicas, { id: item.id, text: item.text, done: false }] }));
    setNewDinSporadic("");
  }
  function removeDinFromSporadicDraft(id) {
    setSporadicDraft((d) => ({ ...d, dinamicas: d.dinamicas.filter((x) => x.id !== id) }));
  }
  function addEscToSporadicDraft() {
    const text = newEscSporadic.trim();
    if (!text) return;
    const existingCat = (roster.escenasCatalog || []).find((c) => c.text.toLowerCase() === text.toLowerCase());
    const item = existingCat || addEscena(text);
    if (!item || sporadicDraft.escenas.some((d) => d.id === item.id)) { setNewEscSporadic(""); return; }
    setSporadicDraft((d) => ({ ...d, escenas: [...d.escenas, { id: item.id, text: item.text, done: false }] }));
    setNewEscSporadic("");
  }
  function removeEscFromSporadicDraft(id) {
    setSporadicDraft((d) => ({ ...d, escenas: d.escenas.filter((x) => x.id !== id) }));
  }
  function createSporadicClass() {
    const nombre = sporadicDraft.nombre.trim();
    if (!nombre) return;
    const newGroup = { id: `suelta-${uid()}`, label: nombre, level: nombre, day: null, dow: null };
    const newSession = {
      id: uid(), groupId: newGroup.id, date: pickerDate,
      profesor: sporadicDraft.profesor, tipoClase: "",
      estado: sporadicDraft.estado,
      attendance: {}, dinamicas: sporadicDraft.dinamicas,
      escenas: sporadicDraft.escenas, titulos: sporadicDraft.titulos,
      conceptosFinales: sporadicDraft.conceptosFinales, observaciones: sporadicDraft.observaciones,
      updatedAt: Date.now(),
    };
    persistRoster({ ...roster, clasesSueltas: [...(roster.clasesSueltas || []), newGroup], students: { ...roster.students, [newGroup.id]: [] } });
    persistSessions([...sessions, newSession]);
    setPickerDate(null);
    setCreatingSporadic(false);
    resetSporadicDraft();
    goToAlumnosGroup(newGroup.id);
  }

  return (
    <div>
      <Card>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
          <button style={btn("ghost", { padding: "6px 10px" })} onClick={() => setCursor(new Date(cyear, cmonth - 1, 1))}><ChevronLeft size={16} /></button>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ fontSize: 17, fontWeight: 800, textTransform: "capitalize" }}>{MONTH_NAMES[cmonth]} {cyear}</div>
          </div>
          <button style={btn("ghost", { padding: "6px 10px" })} onClick={() => setCursor(new Date(cyear, cmonth + 1, 1))}><ChevronRight size={16} /></button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4, marginBottom: 4 }}>
          {DOW_NAMES.map((d) => <div key={d} className="cal-dow" style={{ textAlign: "center", fontSize: 11, fontWeight: 700, color: T.muted, padding: "4px 0" }}>{d}</div>)}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4 }}>
          {cells.map((d, i) => {
            if (d === null) return <div key={i} />;
            const dateStr = `${cyear}-${pad2(cmonth + 1)}-${pad2(d)}`;
            let sess = byDate[dateStr] || [];
            if (!isAdmin) sess = sess.filter((s) => s.profesor === auth.profesorName);
            const isToday = dateStr === iso(new Date());
            const dow = new Date(cyear, cmonth, d).getDay();
            const dayGroupsCount = isAdmin ? GROUPS.filter((g) => g.dow === dow).length : sess.length;
            return (
              <button key={i} className="cal-cell" onClick={() => handleDayClick(dateStr)} style={{
                minHeight: 70, borderRadius: 8, border: `1px solid ${isToday ? T.accent : T.paperDim}`,
                background: isToday ? T.accentDim : "#fff", padding: 0, overflow: "hidden",
                textAlign: "left", display: "flex", flexDirection: "column",
                opacity: dayGroupsCount > 0 || sess.length ? 1 : 0.45,
              }}>
                <div className="cal-day-num" style={{ fontSize: 11, color: T.muted, padding: "4px 6px" }}>{d}</div>
                <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
                  {sess.map((s) => {
                    const given = isGiven(s);
                    const g = findGroup(roster, s.groupId);
                    const c = teacherColor(s.profesor);
                    return (
                      <div key={s.id} className="cal-init" title={`${s.profesor || "sin asignar"} — ${g?.label || ""}${given ? " (impartida)" : " (programada)"}`} style={{
                        flex: 1, minHeight: 18, background: c,
                        color: given ? T.ink : "#fff",
                        fontSize: 9, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center",
                        borderTop: `1px solid rgba(255,255,255,0.35)`, textAlign: "center", lineHeight: 1.15, padding: "2px 1px",
                      }}>{levelAbbrev(g?.level)} {s.profesor ? s.profesor.trim()[0].toUpperCase() : "?"}</div>
                    );
                  })}
                </div>
              </button>
            );
          })}
        </div>
        <div className="cal-legend" style={{ display: "flex", gap: 14, marginTop: 12, fontSize: 11.5, color: T.muted, flexWrap: "wrap" }}>
          <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 12, height: 12, borderRadius: 3, background: T.muted, color: T.ink, fontSize: 8, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center" }}>A</span> Impartida (letra oscura)</span>
          <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 12, height: 12, borderRadius: 3, background: T.muted, color: "#fff", fontSize: 8, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center" }}>A</span> Programada (letra blanca)</span>
        </div>
      </Card>

      {isAdmin && (
        <Card>
          <Label>Leyenda de profesores</Label>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
            {roster.teachers.map((t) => (
              <span key={t.name} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13 }}>
                <span style={{ width: 16, height: 16, borderRadius: 4, background: t.color, color: "#fff", fontSize: 9, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center" }}>{t.name.trim()[0]?.toUpperCase()}</span>
                {t.name}
              </span>
            ))}
            {roster.teachers.length === 0 && <span style={{ fontSize: 13, color: T.muted }}>Añade profesores en la pestaña "Profesores".</span>}
          </div>
        </Card>
      )}

      {toast && (
        <div style={{ position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)", background: T.ink, color: "#fff", padding: "10px 16px", borderRadius: 8, fontSize: 13, zIndex: 80 }}>{toast}</div>
      )}

      {pickerDate && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }} onClick={() => { setPickerDate(null); setCreatingSporadic(false); resetSporadicDraft(); }}>
          <div className="modal-sheet-pad" style={{ background: "#fff", borderRadius: 16, width: "100%", maxWidth: 420, maxHeight: "85vh", overflowY: "auto", padding: 20 }} onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div style={{ fontWeight: 800, fontSize: 15, textTransform: "capitalize" }}>{fmtDateLong(pickerDate)}</div>
              <button onClick={() => { setPickerDate(null); setCreatingSporadic(false); resetSporadicDraft(); }} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={18} /></button>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {dayGroupsFor(pickerDate).map((g) => {
                const s = (byDate[pickerDate] || []).find((x) => x.groupId === g.id);
                const given = isGiven(s);
                const c = s?.profesor ? teacherColor(s.profesor) : null;
                return (
                  <button key={g.id} onClick={() => { setOpenSession({ date: pickerDate, groupId: g.id }); setPickerDate(null); }} style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between", padding: 13, borderRadius: 10,
                    border: `1px solid ${c || T.line}`, background: c ? (given ? c : "#fff") : "#fff", color: c ? (given ? "#fff" : T.ink) : T.ink,
                  }}>
                    <span style={{ fontWeight: 700 }}>{g.level}</span>
                    <span style={{ fontSize: 11.5, opacity: 0.85 }}>{s ? `${s.profesor || "sin asignar"} · ${given ? "impartida" : "programada"}` : "sin organizar"}</span>
                  </button>
                );
              })}
              {(byDate[pickerDate] || []).filter((s) => !GROUPS.some((g) => g.id === s.groupId)).map((s) => {
                const g = findGroup(roster, s.groupId);
                if (!g) return null;
                const given = isGiven(s);
                const c = s.profesor ? teacherColor(s.profesor) : null;
                return (
                  <button key={s.id} onClick={() => { setOpenSession({ date: pickerDate, groupId: g.id }); setPickerDate(null); }} style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between", padding: 13, borderRadius: 10,
                    border: `1px solid ${c || T.line}`, background: c ? (given ? c : "#fff") : "#fff", color: c ? (given ? "#fff" : T.ink) : T.ink,
                  }}>
                    <span style={{ fontWeight: 700 }}>{g.level}</span>
                    <span style={{ fontSize: 11.5, opacity: 0.85 }}>{s.profesor || "sin asignar"} · {given ? "impartida" : "programada"}</span>
                  </button>
                );
              })}
            </div>

            {isAdmin && !creatingSporadic && (
              <button style={btn("ghost", { width: "100%", justifyContent: "center", marginTop: 10 })} onClick={() => setCreatingSporadic(true)}>
                <Plus size={15} /> Crear clase nueva
              </button>
            )}
            {isAdmin && creatingSporadic && (
              <div style={{ marginTop: 12, borderTop: `1px solid ${T.paperDim}`, paddingTop: 12 }}>
                <Label>Nombre del grupo</Label>
                <input autoFocus style={inputStyle({ marginBottom: 10 })} placeholder="Ej. Voz individual, Taller de clown…" value={sporadicDraft.nombre} onChange={(e) => setSporadicDraft((d) => ({ ...d, nombre: e.target.value }))} />
                <Label>Profesor/a</Label>
                <select style={inputStyle({ marginBottom: 10 })} value={sporadicDraft.profesor} onChange={(e) => setSporadicDraft((d) => ({ ...d, profesor: e.target.value }))}>
                  <option value="">Sin asignar</option>
                  {roster.teachers.map((t) => <option key={t.name} value={t.name}>{t.name}</option>)}
                </select>
                <Label>Estado</Label>
                <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
                  <button style={btn(sporadicDraft.estado === "programada" ? "primary" : "ghost", { flex: 1, justifyContent: "center", fontSize: 12.5 })} onClick={() => setSporadicDraft((d) => ({ ...d, estado: "programada" }))}>Programada</button>
                  <button style={btn(sporadicDraft.estado === "impartida" ? "primary" : "ghost", { flex: 1, justifyContent: "center", fontSize: 12.5 })} onClick={() => setSporadicDraft((d) => ({ ...d, estado: "impartida" }))}>Impartida</button>
                </div>

                <Label>Dinámicas</Label>
                {sporadicDraft.dinamicas.map((d) => (
                  <div key={d.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", fontSize: 13 }}>
                    <span>{d.text}</span>
                    <button onClick={() => removeDinFromSporadicDraft(d.id)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Trash2 size={13} /></button>
                  </div>
                ))}
                <div style={{ display: "flex", gap: 8, margin: "4px 0 10px" }}>
                  <input style={inputStyle()} placeholder="Añadir dinámica (nueva o del catálogo)" value={newDinSporadic} onChange={(e) => setNewDinSporadic(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addDinToSporadicDraft()} list="dinamicas-catalog-list-suelta" />
                  <datalist id="dinamicas-catalog-list-suelta">{roster.dinamicasCatalog.map((c) => <option key={c.id} value={c.text} />)}</datalist>
                  <button style={btn("ghost")} onClick={addDinToSporadicDraft}><Plus size={15} /></button>
                </div>

                <Label>Escenas</Label>
                {sporadicDraft.escenas.map((d) => (
                  <div key={d.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", fontSize: 13 }}>
                    <span>{d.text}</span>
                    <button onClick={() => removeEscFromSporadicDraft(d.id)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Trash2 size={13} /></button>
                  </div>
                ))}
                <div style={{ display: "flex", gap: 8, margin: "4px 0 10px" }}>
                  <input style={inputStyle()} placeholder="Añadir escena (nueva o del catálogo)" value={newEscSporadic} onChange={(e) => setNewEscSporadic(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addEscToSporadicDraft()} list="escenas-catalog-list-suelta" />
                  <datalist id="escenas-catalog-list-suelta">{(roster.escenasCatalog || []).map((c) => <option key={c.id} value={c.text} />)}</datalist>
                  <button style={btn("ghost")} onClick={addEscToSporadicDraft}><Plus size={15} /></button>
                </div>

                <Label>Títulos</Label>
                <textarea style={inputStyle({ minHeight: 70, resize: "vertical", marginBottom: 10 })} value={sporadicDraft.titulos} onChange={(e) => setSporadicDraft((d) => ({ ...d, titulos: e.target.value }))} />
                <Label>Conceptos finales</Label>
                <textarea style={inputStyle({ minHeight: 70, resize: "vertical", marginBottom: 10 })} value={sporadicDraft.conceptosFinales} onChange={(e) => setSporadicDraft((d) => ({ ...d, conceptosFinales: e.target.value }))} />
                <Label>Notas</Label>
                <textarea style={inputStyle({ minHeight: 70, resize: "vertical", marginBottom: 14 })} value={sporadicDraft.observaciones} onChange={(e) => setSporadicDraft((d) => ({ ...d, observaciones: e.target.value }))} />

                <div style={{ display: "flex", gap: 8 }}>
                  <button style={btn("ghost", { flex: 1, justifyContent: "center" })} onClick={() => { setCreatingSporadic(false); resetSporadicDraft(); }}>Cancelar</button>
                  <button style={btn("primary", { flex: 1, justifyContent: "center" })} onClick={createSporadicClass}>Crear clase</button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {openSession && (
        <ClassEditor date={openSession.date} groupId={openSession.groupId} ctx={ctx} onClose={() => setOpenSession(null)} />
      )}
    </div>
  );
}

/* ============ FICHA DE CLASE (Calendario + Pasar lista fusionados) ============ */
function ClassEditor({ date, groupId, ctx, onClose }) {
  const { roster, sessions, persistSessions, persistRoster, pending, persistPending, addDinamica, addEscena, isAdmin, auth, askConfirm } = ctx;
  const group = findGroup(roster, groupId);
  const isSporadic = group && !GROUPS.some((g) => g.id === groupId);
  const students = group ? (roster.students[groupId] || []).filter((s) => s.active !== false) : [];
  const existing = sessions.find((s) => s.groupId === groupId && s.date === date);

  const pendingIds = useMemo(() => getPending(pending, groupId), [pending, groupId]);

  const [draft, setDraft] = useState(() => {
    const baseAttendance = {};
    students.forEach((s) => { baseAttendance[s.id] = existing?.attendance?.[s.id] ?? true; });
    let dinamicas = existing?.dinamicas || null;
    if (!dinamicas) {
      dinamicas = pendingIds.map((id) => {
        const cat = roster.dinamicasCatalog.find((c) => c.id === id);
        return { id, text: cat ? cat.text : "(dinámica eliminada)", done: false };
      });
    }
    return {
      profesor: existing?.profesor || (auth.role === "profesor" ? auth.profesorName : ""),
      tipoClase: existing?.tipoClase || "",
      estado: existing?.estado || (date <= iso(new Date()) ? "impartida" : "programada"),
      attendance: baseAttendance,
      dinamicas,
      escenas: Array.isArray(existing?.escenas) ? existing.escenas : [],
      titulos: existing?.titulos || "",
      conceptosFinales: existing?.conceptosFinales || "",
      observaciones: existing?.observaciones || "",
    };
  });
  const [wasPendingIds] = useState(() => new Set(pendingIds));
  const [sub, setSub] = useState("clase");
  const [newDinamica, setNewDinamica] = useState("");
  const [newEscena, setNewEscena] = useState("");
  const [newParticipant, setNewParticipant] = useState("");
  const [saved, setSaved] = useState(false);

  if (!group) {
    // Sesión "huérfana": su grupo ya no existe (p.ej. tras renombrar los grupos). La mostramos
    // igualmente para poder verla y eliminarla, en vez de esconderla u ocultar el error.
    return (
      <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
        <div style={{ background: "#fff", borderRadius: "16px 16px 0 0", width: "100%", maxWidth: 480, padding: 20 }} onClick={(e) => e.stopPropagation()}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div style={{ fontWeight: 800, fontSize: 16, textTransform: "capitalize" }}>{fmtDateLong(date)}</div>
            <button onClick={onClose} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={20} /></button>
          </div>
          <div style={{ fontSize: 13.5, color: T.muted, marginBottom: 16 }}>
            Esta clase pertenece a un grupo que ya no existe (probablemente por un cambio de nombre de grupo). Profesor: <strong>{existing?.profesor || "—"}</strong>.
          </div>
          {isAdmin && existing && (
            <button style={btn("danger", { width: "100%", justifyContent: "center" })} onClick={() => askConfirm("¿Eliminar esta clase?", async () => { await persistSessions(sessions.filter((s) => s.id !== existing.id)); onClose(); })}>
              <Trash2 size={14} /> Eliminar esta clase
            </button>
          )}
        </div>
      </div>
    );
  }

  function patch(p) { setDraft({ ...draft, ...p }); }
  function toggleAttendance(studentId) {
    patch({ attendance: { ...draft.attendance, [studentId]: !draft.attendance[studentId] } });
  }
  function toggleDinamica(id) {
    patch({ dinamicas: draft.dinamicas.map((d) => (d.id === id ? { ...d, done: !d.done } : d)) });
  }
  function addDinamicaToDraft() {
    const text = newDinamica.trim();
    if (!text) return;
    const existingCat = roster.dinamicasCatalog.find((c) => c.text.toLowerCase() === text.toLowerCase());
    const item = existingCat || addDinamica(text);
    if (!item) return;
    if (draft.dinamicas.some((d) => d.id === item.id)) { setNewDinamica(""); return; }
    patch({ dinamicas: [...draft.dinamicas, { id: item.id, text: item.text, done: false }] });
    setNewDinamica("");
  }
  function removeDinamicaFromDraft(id) {
    patch({ dinamicas: draft.dinamicas.filter((d) => d.id !== id) });
  }
  function toggleEscena(id) {
    patch({ escenas: draft.escenas.map((d) => (d.id === id ? { ...d, done: !d.done } : d)) });
  }
  function addEscenaToDraft() {
    const text = newEscena.trim();
    if (!text) return;
    const existingCat = (roster.escenasCatalog || []).find((c) => c.text.toLowerCase() === text.toLowerCase());
    const item = existingCat || addEscena(text);
    if (!item) return;
    if (draft.escenas.some((d) => d.id === item.id)) { setNewEscena(""); return; }
    patch({ escenas: [...draft.escenas, { id: item.id, text: item.text, done: false }] });
    setNewEscena("");
  }
  function removeEscenaFromDraft(id) {
    patch({ escenas: draft.escenas.filter((d) => d.id !== id) });
  }
  function addParticipant() {
    const name = newParticipant.trim();
    if (!name) return;
    const newS = { id: uid(), name, active: true };
    const nextList = [...(roster.students[groupId] || []), newS];
    persistRoster({ ...roster, students: { ...roster.students, [groupId]: nextList } });
    patch({ attendance: { ...draft.attendance, [newS.id]: true } });
    setNewParticipant("");
  }
  async function deletePendingForever(id) {
    askConfirm("¿Quitar esta dinámica pendiente definitivamente? Dejará de aparecer en las próximas clases de este grupo.", async () => {
      await persistPending(removeFromPending(pending, groupId, id));
      removeDinamicaFromDraft(id);
    });
  }
  // Los tipos de clase (con sus dinámicas y escenas) son específicos de cada nivel: los de
  // Iniciación 1 no deben poder aplicarse en Avanzado, Veteranos, etc., para que las dinámicas
  // de un nivel no se crucen ni contaminen las de otro. Los tipos antiguos sin "nivel" guardado
  // se tratan como de Iniciación 1 (el único nivel que tenía catálogo cuando se creó esto).
  const tiposDeEsteGrupo = useMemo(
    () => roster.tiposClase.filter((t) => (t.nivel || "Iniciación 1") === group.level).sort((a, b) => a.numero - b.numero),
    [roster.tiposClase, group.level]
  );
  function applyTipoClase(numeroStr) {
    const numero = parseInt(numeroStr, 10);
    const tipo = tiposDeEsteGrupo.find((t) => t.numero === numero);
    if (!tipo) { patch({ tipoClase: numeroStr }); return; }
    setDraft((d) => {
      // Las dinámicas propias de este tipo de clase van primero (recién marcadas sin hacer);
      // lo que ya hubiera en el borrador (típicamente pendientes de clases anteriores de este
      // grupo) se conserva detrás, como "extra" a intentar si sobra tiempo.
      const tipoItems = (tipo.dinamicaIds || []).map((id) => {
        const cat = roster.dinamicasCatalog.find((c) => c.id === id);
        const already = d.dinamicas.find((x) => x.id === id);
        return { id, text: cat ? cat.text : "(dinámica eliminada)", done: already ? already.done : false };
      });
      const tipoIdSet = new Set(tipoItems.map((x) => x.id));
      const extras = d.dinamicas.filter((x) => !tipoIdSet.has(x.id));
      // Las escenas de la plantilla del tipo de clase estaban guardadas como texto libre
      // (una por línea); las convertimos en escenas fijas con check, enlazándolas con el
      // catálogo de escenas (o creándolas si es la primera vez que se usan).
      let escenas = d.escenas;
      if (!escenas.length && tipo.escenas) {
        escenas = tipo.escenas.split("\n").map((l) => l.trim()).filter(Boolean).map((text) => {
          const cat = (roster.escenasCatalog || []).find((c) => c.text.toLowerCase() === text.toLowerCase());
          const item = cat || addEscena(text);
          return item ? { id: item.id, text: item.text, done: false } : null;
        }).filter(Boolean);
      }
      return { ...d, tipoClase: numeroStr, dinamicas: [...tipoItems, ...extras], escenas };
    });
  }

  async function save() {
    const session = {
      id: existing?.id || uid(), groupId, date,
      profesor: draft.profesor, tipoClase: draft.tipoClase, estado: draft.estado,
      attendance: draft.attendance, dinamicas: draft.dinamicas,
      escenas: draft.escenas, titulos: draft.titulos, conceptosFinales: draft.conceptosFinales, observaciones: draft.observaciones,
      updatedAt: Date.now(),
    };
    const others = sessions.filter((s) => !(s.groupId === groupId && s.date === date));
    await persistSessions([...others, session]);
    // Las dinámicas solo pasan a "pendientes" cuando la clase realmente se ha impartido — si solo
    // está programada, marcar algo como "no hecho" no significa nada todavía.
    if (draft.estado === "impartida") {
      await persistPending(applyDynamicsToPending(pending, groupId, draft.dinamicas));
    }
    setSaved(true);
    setTimeout(() => setSaved(false), 1800);
  }
  async function exportWord() {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } = await import("docx");
    const thinBorder = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
    const cellBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
    const checklistRows = (items, emptyText) => {
      if (!items || items.length === 0) return [new Paragraph({ text: emptyText, italics: true })];
      return items.map((it) => new Paragraph({ text: `☐  ${it.text}`, spacing: { after: 80 } }));
    };
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({ text: "LABORATORIO IMPRO", heading: HeadingLevel.HEADING_3, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: group.label, heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
          new Paragraph({ text: fmtDateLong(date), alignment: AlignmentType.CENTER, spacing: { after: 200 } }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({ children: [
                new TableCell({ borders: cellBorders, width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Profesor/a: ", bold: true }), new TextRun(draft.profesor || "—")] })] }),
                new TableCell({ borders: cellBorders, width: { size: 50, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Estado: ", bold: true }), new TextRun(given ? "Impartida" : "Programada")] })] }),
              ] }),
            ],
          }),
          new Paragraph({ text: "", spacing: { after: 200 } }),
          new Paragraph({ text: "Alumnado", heading: HeadingLevel.HEADING_2 }),
          ...(students.length === 0
            ? [new Paragraph({ text: "Sin alumnos en este grupo.", italics: true })]
            : students.map((s) => new Paragraph({ text: `☐  ${s.name}`, spacing: { after: 60 } }))),
          new Paragraph({ text: "", spacing: { after: 200 } }),
          new Paragraph({ text: "Dinámicas", heading: HeadingLevel.HEADING_2 }),
          ...checklistRows(draft.dinamicas, "Sin dinámicas."),
          new Paragraph({ text: "", spacing: { after: 200 } }),
          new Paragraph({ text: "Escenas", heading: HeadingLevel.HEADING_2 }),
          ...checklistRows(draft.escenas, "Sin escenas."),
          new Paragraph({ text: "", spacing: { after: 200 } }),
          new Paragraph({ text: "Títulos", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: draft.titulos || "—" }),
          new Paragraph({ text: "", spacing: { after: 200 } }),
          new Paragraph({ text: "Conceptos finales", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: draft.conceptosFinales || "—" }),
          new Paragraph({ text: "", spacing: { after: 200 } }),
          new Paragraph({ text: "Notas", heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: draft.observaciones || "—" }),
        ],
      }],
    });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${group.label} - ${date}.docx`.replace(/[/\\?%*:|"<>]/g, "-");
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }
  async function deleteSession() {
    if (!existing) { onClose(); return; }
    askConfirm("¿Eliminar esta clase de este día?", async () => {
      await persistSessions(sessions.filter((s) => s.id !== existing.id));
      onClose();
    });
  }

  // Ambos roles ven todas las subpestañas; lo que cambia es qué puede tocar cada uno dentro de ellas.
  const subTabs = [
    { id: "resumen", label: "Resumen" },
    { id: "clase", label: isAdmin ? "Clase" : "Asistencia" },
    { id: "dinamicas", label: "Dinámicas" },
    { id: "escenas", label: "Escenas" },
    { id: "titulos", label: "Títulos" },
    { id: "conceptos", label: "Conceptos finales" },
    { id: "notas", label: "Notas" },
  ];
  const given = draft.estado === "impartida";

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 50 }} onClick={onClose}>
      <div className="modal-sheet" style={{ background: "#fff", borderRadius: "16px 16px 0 0", width: "100%", maxWidth: 620, maxHeight: "90vh", overflowY: "auto", padding: "20px 20px 0" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
          <div>
            <div style={{ fontWeight: 800, fontSize: 16, textTransform: "capitalize" }}>{fmtDateLong(date)}</div>
            <div style={{ fontSize: 12.5, color: T.muted, display: "flex", alignItems: "center", gap: 6 }}>
              {group.label}
              {isSporadic && <Badge tone="blue">Clase suelta</Badge>}
            </div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={20} /></button>
        </div>
        {existing && (
          <Badge tone={given ? "green" : "accent"}>{given ? "Impartida" : "Programada"}</Badge>
        )}

        <div style={{ marginTop: 12 }}>
          <SubTabs tabs={subTabs} active={sub} setActive={setSub} />
        </div>

        {sub === "resumen" && (
          <div>
            <div style={{ marginBottom: 16 }}>
              <Label>Dinámicas</Label>
              {draft.dinamicas.length === 0 ? (
                <div style={{ fontSize: 13, color: T.muted }}>Sin dinámicas.</div>
              ) : draft.dinamicas.map((d) => (
                <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, padding: "3px 0", color: d.done ? T.muted : T.ink, textDecoration: d.done ? "line-through" : "none" }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: d.done ? T.green : T.accent, flexShrink: 0 }} />
                  {d.text}
                </div>
              ))}
            </div>
            <div style={{ marginBottom: 16 }}>
              <Label>Escenas</Label>
              {draft.escenas.length === 0 ? (
                <div style={{ fontSize: 13, color: T.muted }}>Sin escenas.</div>
              ) : draft.escenas.map((d) => (
                <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, padding: "3px 0", color: d.done ? T.muted : T.ink, textDecoration: d.done ? "line-through" : "none" }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: d.done ? T.green : T.accent, flexShrink: 0 }} />
                  {d.text}
                </div>
              ))}
            </div>
            <div style={{ marginBottom: 16 }}>
              <Label>Títulos</Label>
              <div style={{ fontSize: 13.5, whiteSpace: "pre-wrap", color: draft.titulos ? T.ink : T.muted }}>{draft.titulos || "Sin títulos."}</div>
            </div>
            <div>
              <Label>Conceptos finales</Label>
              <div style={{ fontSize: 13.5, whiteSpace: "pre-wrap", color: draft.conceptosFinales ? T.ink : T.muted }}>{draft.conceptosFinales || "Sin conceptos finales."}</div>
            </div>
          </div>
        )}

        {sub === "clase" && (
          <div>
            {isAdmin && (
              <div className="two-col" style={{ gap: 10, marginBottom: 8 }}>
                <div>
                  <Label>Profesor/a</Label>
                  <select style={inputStyle()} value={draft.profesor} onChange={(e) => patch({ profesor: e.target.value })}>
                    <option value="">Sin asignar</option>
                    {roster.teachers.map((t) => <option key={t.name} value={t.name}>{t.name}</option>)}
                  </select>
                </div>
                {tiposDeEsteGrupo.length > 0 && (
                  <div>
                    <Label>Tipo de clase</Label>
                    <select style={inputStyle()} value={draft.tipoClase} onChange={(e) => applyTipoClase(e.target.value)}>
                      <option value="">Sin asignar</option>
                      {tiposDeEsteGrupo.map((tipo) => (
                        <option key={tipo.numero} value={tipo.numero}>{tipo.numero}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
            )}
            {isAdmin && (
              <div style={{ marginBottom: 14 }}>
                <Label>Estado</Label>
                <div style={{ display: "flex", gap: 6 }}>
                  <button style={btn(draft.estado === "programada" ? "primary" : "ghost", { flex: 1, justifyContent: "center", fontSize: 12.5, padding: "8px 10px" })} onClick={() => patch({ estado: "programada" })}>Programada</button>
                  <button style={btn(draft.estado === "impartida" ? "primary" : "ghost", { flex: 1, justifyContent: "center", fontSize: 12.5, padding: "8px 10px" })} onClick={() => patch({ estado: "impartida" })}>Impartida</button>
                </div>
              </div>
            )}
            {isAdmin && draft.tipoClase && tiposDeEsteGrupo.length > 0 && (
              <button style={btn("primary", { fontSize: 13, padding: "9px 14px", marginBottom: 14, background: T.accent })} onClick={() => applyTipoClase(draft.tipoClase)}><Sparkles size={14} /> Cargar clase</button>
            )}
            {students.length === 0 ? (
              <div style={{ fontSize: 13, color: T.muted, marginBottom: 10 }}>{isSporadic ? "Añade los nombres de los participantes de esta clase abajo." : "Este grupo no tiene alumnos activos todavía."}</div>
            ) : (
              <div>
                <Label>Alumnos — {students.length}</Label>
                {students.map((s) => {
                  const present = draft.attendance[s.id];
                  return (
                    <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderTop: `1px solid ${T.paperDim}` }}>
                      <Avatar name={s.name} photo={s.photo} size={30} />
                      <div style={{ flex: 1, fontSize: 14, color: present ? T.ink : T.muted }}>{s.name}</div>
                      <button onClick={() => toggleAttendance(s.id)} style={attendanceBoxStyle(present, 24)}>{present ? <Check size={15} color="#fff" /> : null}</button>
                    </div>
                  );
                })}
              </div>
            )}
            {isAdmin && isSporadic && (
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <input style={inputStyle()} placeholder="Nombre del participante" value={newParticipant} onChange={(e) => setNewParticipant(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addParticipant()} />
                <button style={btn("ghost")} onClick={addParticipant}><Plus size={15} /></button>
              </div>
            )}
          </div>
        )}

        {sub === "dinamicas" && (
          <div>
            {draft.dinamicas.length === 0 && <div style={{ fontSize: 13, color: T.muted, marginBottom: 10 }}>Sin dinámicas en esta clase todavía.{isAdmin ? " Selecciona un tipo de clase en \"Clase\" y carga su plantilla, o añade dinámicas sueltas abajo." : ""}</div>}
            {draft.dinamicas.map((d) => {
              const isPending = wasPendingIds.has(d.id);
              return (
                <div key={d.id} style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "8px 10px", borderRadius: 8, marginBottom: 6,
                  background: isPending ? T.greenBg : T.paper, border: `1px solid ${isPending ? T.green : T.paperDim}`,
                }}>
                  <button onClick={() => toggleDinamica(d.id)} style={attendanceBoxStyle(d.done, 22)}>{d.done ? <Check size={14} color="#fff" /> : null}</button>
                  <div style={{ flex: 1, fontSize: 13.5, color: isPending && !d.done ? T.green : T.ink, fontWeight: isPending ? 700 : 500 }}>
                    {d.text}{isPending && !d.done && <span style={{ fontSize: 10.5, fontWeight: 700, marginLeft: 6, textTransform: "uppercase" }}>· pendiente</span>}
                  </div>
                  {isAdmin && isPending && (
                    <button onClick={() => deletePendingForever(d.id)} title="Quitar de pendientes definitivamente" style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={14} /></button>
                  )}
                  {isAdmin && !isPending && (
                    <button onClick={() => removeDinamicaFromDraft(d.id)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Trash2 size={13} /></button>
                  )}
                </div>
              );
            })}
            {isAdmin && (
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <input style={inputStyle()} placeholder="Añadir dinámica (nueva o del catálogo)" value={newDinamica} onChange={(e) => setNewDinamica(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addDinamicaToDraft()} list="dinamicas-catalog-list" />
                <datalist id="dinamicas-catalog-list">
                  {roster.dinamicasCatalog.map((c) => <option key={c.id} value={c.text} />)}
                </datalist>
                <button style={btn("ghost")} onClick={addDinamicaToDraft}><Plus size={15} /></button>
              </div>
            )}
          </div>
        )}

        {sub === "escenas" && (
          <div>
            {draft.escenas.length === 0 && <div style={{ fontSize: 13, color: T.muted, marginBottom: 10 }}>Sin escenas en esta clase todavía.</div>}
            {draft.escenas.map((d) => (
              <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 10px", borderRadius: 8, marginBottom: 6, background: T.paper, border: `1px solid ${T.paperDim}` }}>
                <button onClick={() => toggleEscena(d.id)} style={attendanceBoxStyle(d.done, 22)}>{d.done ? <Check size={14} color="#fff" /> : null}</button>
                <div style={{ flex: 1, fontSize: 13.5, color: T.ink }}>{d.text}</div>
                {isAdmin && (
                  <button onClick={() => removeEscenaFromDraft(d.id)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Trash2 size={13} /></button>
                )}
              </div>
            ))}
            {isAdmin && (
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <input style={inputStyle()} placeholder="Añadir escena (nueva o del catálogo)" value={newEscena} onChange={(e) => setNewEscena(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addEscenaToDraft()} list="escenas-catalog-list" />
                <datalist id="escenas-catalog-list">
                  {(roster.escenasCatalog || []).map((c) => <option key={c.id} value={c.text} />)}
                </datalist>
                <button style={btn("ghost")} onClick={addEscenaToDraft}><Plus size={15} /></button>
              </div>
            )}
          </div>
        )}
        {sub === "titulos" && (
          <div><Label>Títulos</Label><textarea style={inputStyle({ minHeight: 100, resize: "vertical" })} value={draft.titulos} onChange={(e) => patch({ titulos: e.target.value })} placeholder="Títulos que han surgido durante la clase…" /></div>
        )}
        {sub === "conceptos" && (
          <div><Label>Conceptos finales</Label><textarea style={inputStyle({ minHeight: 100, resize: "vertical" })} value={draft.conceptosFinales} onChange={(e) => patch({ conceptosFinales: e.target.value })} placeholder="Conceptos que se han trabajado…" /></div>
        )}
        {sub === "notas" && (
          <div><Label>Notas</Label><textarea style={inputStyle({ minHeight: 100, resize: "vertical" })} value={draft.observaciones} onChange={(e) => patch({ observaciones: e.target.value })} placeholder="Cómo ha ido la clase, incidencias…" /></div>
        )}

        <div className="modal-sheet-pad" style={{ display: "flex", gap: 8, marginTop: 18, borderTop: `1px solid ${T.paperDim}`, padding: "14px 0", background: "#fff", position: "sticky", bottom: 0 }}>
          <button style={btn("primary", { flex: 1, justifyContent: "center" })} onClick={save}>{saved ? <><Check size={15} /> Guardado</> : <><Save size={15} /> Guardar</>}</button>
          <button style={btn("ghost")} title="Descargar esta clase en Word, para imprimirla en papel" onClick={exportWord}><FileDown size={15} /></button>
          {isAdmin && existing && <button style={btn("danger")} onClick={deleteSession}><Trash2 size={14} /></button>}
        </div>
      </div>
    </div>
  );
}

/* ============ ALUMNOS (antes "Histórico") ============ */
function GroupPicker({ groupId, setGroupId, roster }) {
  const clasesSueltas = roster?.clasesSueltas || [];
  const [day, setDay] = useState(() => {
    const fixed = GROUPS.find((g) => g.id === groupId);
    if (fixed) return fixed.day;
    if (clasesSueltas.some((g) => g.id === groupId)) return "extra";
    return DAYS[0].key;
  });
  const dayGroups = day === "extra" ? clasesSueltas : GROUPS.filter((g) => g.day === day);
  useEffect(() => {
    if (dayGroups.length && !dayGroups.some((g) => g.id === groupId)) setGroupId(dayGroups[0].id);
    // eslint-disable-next-line
  }, [day]);
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
        {DAYS.map((d) => (
          <button key={d.key} onClick={() => setDay(d.key)} style={{
            padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 600,
            border: `1px solid ${day === d.key ? T.ink : T.line}`,
            background: day === d.key ? T.ink : "#fff", color: day === d.key ? "#fff" : T.ink,
          }}>{d.label}</button>
        ))}
        {roster && (
          <button onClick={() => setDay("extra")} style={{
            padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 600,
            border: `1px solid ${day === "extra" ? T.ink : T.line}`,
            background: day === "extra" ? T.ink : "#fff", color: day === "extra" ? "#fff" : T.ink,
          }}>Clases extra</button>
        )}
      </div>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {day === "extra" && dayGroups.length === 0 && (
          <div style={{ fontSize: 12.5, color: T.muted }}>Aún no hay clases sueltas creadas.</div>
        )}
        {dayGroups.map((g) => (
          <button key={g.id} onClick={() => setGroupId(g.id)} style={{
            padding: "6px 14px", borderRadius: 999, fontSize: 12.5, fontWeight: 600,
            border: `1px solid ${groupId === g.id ? T.accent : T.line}`,
            background: groupId === g.id ? T.accentDim : "#fff", color: groupId === g.id ? T.accent : T.muted,
          }}>{day === "extra" ? g.label : g.level}</button>
        ))}
      </div>
    </div>
  );
}

function AlumnosTab({ ctx }) {
  const { roster, persistRoster, sessions, persistSessions, payments, focusGroupId, setFocusGroupId } = ctx;
  const [topSub, setTopSub] = useState("lista"); // lista | movimientos
  const [groupId, setGroupId] = useState(GROUPS[0].id);
  const [editing, setEditing] = useState(false);
  const [newName, setNewName] = useState("");
  const [newPhone, setNewPhone] = useState("");
  const [newNotes, setNewNotes] = useState("");
  const [renameId, setRenameId] = useState(null);
  const [renameVal, setRenameVal] = useState("");
  const [search, setSearch] = useState("");
  const [openStudent, setOpenStudent] = useState(null); // {student, group}
  const [bajaFor, setBajaFor] = useState(null);
  const [bajaReason, setBajaReason] = useState("");

  // Cuando se crea una clase suelta nueva desde el Calendario, éste nos pide abrir aquí
  // ese grupo (recién creado, sin alumnos) para poder añadirle alumnos enseguida.
  useEffect(() => {
    if (focusGroupId) {
      setTopSub("lista");
      setGroupId(focusGroupId);
      setEditing(true);
      setFocusGroupId(null);
    }
    // eslint-disable-next-line
  }, [focusGroupId]);

  const group = findGroup(roster, groupId);
  const students = (roster.students[groupId] || []).filter((s) => s.active !== false);

  const searchResults = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return [];
    const out = [];
    Object.entries(roster.students).forEach(([gid, list]) => {
      list.forEach((s) => {
        if (s.active !== false && s.name.toLowerCase().includes(q)) {
          out.push({ student: s, group: findGroup(roster, gid) });
        }
      });
    });
    return out.slice(0, 12);
  }, [search, roster.students]);

  function setStudents(list) { persistRoster({ ...roster, students: { ...roster.students, [groupId]: list } }); }
  function addStudent() {
    const n = newName.trim();
    if (!n) return;
    setStudents([...(roster.students[groupId] || []), { id: uid(), name: n, phone: newPhone.trim(), notes: newNotes.trim(), active: true, fechaAlta: iso(new Date()) }]);
    setNewName(""); setNewPhone(""); setNewNotes("");
  }
  function saveRename() {
    setStudents((roster.students[groupId] || []).map((s) => (s.id === renameId ? { ...s, name: renameVal.trim() || s.name } : s)));
    setRenameId(null);
  }
  function hardDelete(id) {
    setStudents((roster.students[groupId] || []).filter((s) => s.id !== id));
    setBajaFor(null);
  }
  function confirmBaja() {    setStudents((roster.students[groupId] || []).map((s) => (s.id === bajaFor.id ? { ...s, active: false, bajaReason: bajaReason.trim(), bajaDate: iso(new Date()) } : s)));
    setBajaFor(null); setBajaReason("");
  }

  const cols = useMemo(() => weekdayColumnsForRange(group.dow, 2).reverse(), [group.dow]);
  const sessByDate = useMemo(() => {
    const m = {};
    sessions.filter((s) => s.groupId === group.id).forEach((s) => (m[s.date] = s));
    return m;
  }, [sessions, group.id]);
  async function toggleCell(studentId, date) {
    const others = sessions.filter((s) => !(s.groupId === group.id && s.date === date));
    const existing = sessByDate[date];
    const attendance = { ...(existing?.attendance || {}) };
    attendance[studentId] = !(attendance[studentId] ?? true);
    const session = existing ? { ...existing, attendance } : { id: uid(), groupId: group.id, date, profesor: "", tipoClase: "", observaciones: "", dinamicas: [], attendance, updatedAt: Date.now() };
    await persistSessions([...others, session]);
  }

  function exportAsistencia() {
    const wb = XLSX.utils.book_new();
    GROUPS.forEach((g) => {
      const groupStudents = roster.students[g.id] || [];
      if (groupStudents.length === 0) return;
      const groupSessions = sessions.filter((s) => s.groupId === g.id).slice().sort((a, b) => a.date.localeCompare(b.date));
      const dates = groupSessions.map((s) => s.date);
      const header = ["Alumno/a", ...dates.map(fmtDateShort)];
      const rows = groupStudents.map((st) => {
        const row = [st.name];
        groupSessions.forEach((s) => { row.push(s.attendance?.[st.id] ? "X" : ""); });
        return row;
      });
      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      XLSX.utils.book_append_sheet(wb, ws, g.label.slice(0, 31));
    });
    XLSX.writeFile(wb, "Asistencia.xlsx");
  }

  function exportPagos() {
    const now = new Date();
    const months = [];
    for (let i = 11; i >= 0; i--) months.push(yyyymm(new Date(now.getFullYear(), now.getMonth() - i, 1)));
    const header = ["Alumno/a", "Grupo", ...months.map(monthLabel)];
    const rows = [];
    GROUPS.forEach((g) => {
      (roster.students[g.id] || []).forEach((st) => {
        const row = [st.name, g.label];
        months.forEach((ym) => {
          const p = (payments.students[st.id] || {})[ym];
          row.push(p?.paid ? (p.amount ? `${p.amount} €${p.paidTo ? " (" + p.paidTo + ")" : ""}` : "PAGADO") : "");
        });
        rows.push(row);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pagos");
    XLSX.writeFile(wb, "Pagos.xlsx");
  }

  return (
    <div>
      <Card style={{ display: "flex", gap: 6 }}>
        <button style={btn(topSub === "lista" ? "primary" : "ghost", { flex: 1, justifyContent: "center", fontSize: 13 })} onClick={() => setTopSub("lista")}>Alumnos</button>
        <button style={btn(topSub === "movimientos" ? "primary" : "ghost", { flex: 1, justifyContent: "center", fontSize: 13 })} onClick={() => setTopSub("movimientos")}>Movimientos (altas/bajas)</button>
      </Card>

      {topSub === "movimientos" && <MovimientosView roster={roster} />}

      {topSub === "lista" && (<>
      <Card>
        <Label>Buscar alumno/a (todos los grupos)</Label>
        <div style={{ position: "relative" }}>
          <Search size={15} style={{ position: "absolute", left: 11, top: 11, color: T.muted }} />
          <input style={inputStyle({ paddingLeft: 32 })} placeholder="Escribe un nombre…" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
        {searchResults.length > 0 && (
          <div style={{ marginTop: 8, borderTop: `1px solid ${T.paperDim}`, paddingTop: 8 }}>
            {searchResults.map((r) => (
              <button key={r.student.id} onClick={() => setOpenStudent(r)} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", border: "none", background: "transparent", padding: "6px 4px", textAlign: "left", fontSize: 13.5 }}>
                <span style={{ display: "flex", alignItems: "center", gap: 8 }}><Avatar name={r.student.name} photo={r.student.photo} size={26} />{r.student.name}</span><span style={{ color: T.muted, fontSize: 12 }}>{r.group?.label}</span>
              </button>
            ))}
          </div>
        )}
      </Card>

      <Card style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <button style={btn("ghost", { fontSize: 12.5 })} onClick={exportAsistencia}><List size={14} /> Descargar Excel de Asistencia</button>
        <button style={btn("ghost", { fontSize: 12.5 })} onClick={exportPagos}><List size={14} /> Descargar Excel de Pagos</button>
      </Card>

      <GroupPicker groupId={groupId} setGroupId={setGroupId} roster={roster} />

      <Card style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontSize: 13, color: T.muted }}>{group.dow != null ? `Mes actual y los dos anteriores · ${group.label}` : `${group.label} · clase suelta`}</div>
        <button style={btn(editing ? "primary" : "ghost", { padding: "7px 12px", fontSize: 12.5 })} onClick={() => setEditing(!editing)}>
          <Pencil size={13} /> {editing ? "Terminar edición" : "Editar alumnos"}
        </button>
      </Card>

      {editing && (
        <Card>
          <Label>Añadir alumno/a</Label>
          <div className="two-col" style={{ gap: 8, marginBottom: 8 }}>
            <input style={inputStyle()} placeholder="Nombre" value={newName} onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addStudent()} />
            <input style={inputStyle()} placeholder="Teléfono (opcional)" value={newPhone} onChange={(e) => setNewPhone(e.target.value)} />
          </div>
          <input style={inputStyle({ marginBottom: 8 })} placeholder="Comentario / cualidad (opcional)" value={newNotes} onChange={(e) => setNewNotes(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addStudent()} />
          <button style={btn("primary")} onClick={addStudent}><Plus size={15} /> Añadir</button>
        </Card>
      )}

      {students.length === 0 ? (
        <Card style={{ textAlign: "center", color: T.muted }}>Sin alumnos activos en este grupo todavía.</Card>
      ) : group.dow == null ? (
        <Card style={{ padding: 0, overflow: "hidden" }}>
          {students.map((s, i) => (
            <div key={s.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: "10px 14px", borderTop: i ? `1px solid ${T.paperDim}` : "none" }}>
              {renameId === s.id ? (
                <input autoFocus style={inputStyle({ padding: "3px 6px", fontSize: 12.5 })} value={renameVal} onChange={(e) => setRenameVal(e.target.value)} onKeyDown={(e) => e.key === "Enter" && saveRename()} onBlur={saveRename} />
              ) : (
                <button onClick={() => setOpenStudent({ student: s, group })} style={{ display: "flex", alignItems: "center", gap: 8, border: "none", background: "transparent", padding: 0, font: "inherit", color: T.ink, textAlign: "left" }}>
                  <Avatar name={s.name} photo={s.photo} size={28} />
                  <span style={{ textDecoration: "underline", textDecorationColor: T.line, textUnderlineOffset: 3 }}>{s.name}</span>
                </button>
              )}
              {editing && renameId !== s.id && (
                <span style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                  <button onClick={() => { setRenameId(s.id); setRenameVal(s.name); }} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Pencil size={13} /></button>
                  <button onClick={() => setBajaFor(s)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><UserX size={13} /></button>
                </span>
              )}
            </div>
          ))}
        </Card>
      ) : (
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ overflow: "auto", maxHeight: "70vh" }}>
            <table style={{ fontSize: 13 }}>
              <thead>
                <tr>
                  <th style={{ position: "sticky", top: 0, left: 0, zIndex: 3, background: T.ink, color: "#fff", padding: "10px 14px", textAlign: "left", minWidth: 180, borderRight: `1px solid ${T.ink}` }}>Alumno/a</th>
                  {cols.map((c) => (
                    <th key={c} style={{ position: "sticky", top: 0, zIndex: 2, background: T.ink, color: "#fff", padding: "8px 6px", fontWeight: 600, minWidth: 44, borderLeft: `1px solid rgba(255,255,255,0.15)` }}>{fmtDateShort(c)}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {students.map((s, i) => (
                  <tr key={s.id} style={{ background: i % 2 ? T.paperDim : "#fff" }}>
                    <td style={{ position: "sticky", left: 0, zIndex: 1, background: i % 2 ? T.paperDim : "#fff", padding: "8px 14px", borderRight: `1px solid ${T.line}`, borderBottom: `1px solid ${T.paperDim}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 }}>
                      {renameId === s.id ? (
                        <input autoFocus style={inputStyle({ padding: "3px 6px", fontSize: 12.5 })} value={renameVal} onChange={(e) => setRenameVal(e.target.value)} onKeyDown={(e) => e.key === "Enter" && saveRename()} onBlur={saveRename} />
                      ) : (
                        <button onClick={() => setOpenStudent({ student: s, group })} style={{ border: "none", background: "transparent", padding: 0, font: "inherit", color: T.ink, textAlign: "left", textDecoration: "underline", textDecorationColor: T.line, textUnderlineOffset: 3 }}>{s.name}</button>
                      )}
                      {editing && renameId !== s.id && (
                        <span style={{ display: "flex", gap: 4 }}>
                          <button onClick={() => { setRenameId(s.id); setRenameVal(s.name); }} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Pencil size={13} /></button>
                          <button onClick={() => setBajaFor(s)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><UserX size={13} /></button>
                        </span>
                      )}
                    </td>
                    {cols.map((c) => {
                      const present = sessByDate[c]?.attendance?.[s.id];
                      const marked = present !== undefined;
                      return (
                        <td key={c} onClick={() => toggleCell(s.id, c)} style={{ padding: "8px 6px", textAlign: "center", borderBottom: `1px solid ${T.paperDim}`, borderLeft: `1px solid ${T.paperDim}`, cursor: "pointer", background: marked && present ? T.green : "transparent" }}>
                          {marked && present && <Check size={16} color="#fff" style={{ margin: "0 auto", display: "block" }} />}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {bajaFor && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60 }} onClick={() => setBajaFor(null)}>
          <div style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 380, padding: 20 }} onClick={(e) => e.stopPropagation()}>
            <div style={{ fontWeight: 800, fontSize: 15, marginBottom: 4 }}>{bajaFor.name}</div>
            <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 14 }}>¿Eliminar del todo o dar de baja (se conserva el historial)?</div>
            <Label>Motivo (si das de baja)</Label>
            <textarea style={inputStyle({ minHeight: 60, marginBottom: 12 })} value={bajaReason} onChange={(e) => setBajaReason(e.target.value)} placeholder="Opcional" />
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <button style={btn("primary", { justifyContent: "center" })} onClick={confirmBaja}><UserX size={14} /> Dar de baja</button>
              <button style={btn("danger", { justifyContent: "center" })} onClick={() => hardDelete(bajaFor.id)}><Trash2 size={14} /> Eliminar definitivamente</button>
              <button style={btn("ghost", { justifyContent: "center" })} onClick={() => setBajaFor(null)}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {openStudent && (
        <StudentModal student={openStudent.student} group={openStudent.group} ctx={ctx} onClose={() => setOpenStudent(null)}
          onSaveDetails={(patch) => {
            const gid = openStudent.group.id;
            const list = (roster.students[gid] || []).map((s) => (s.id === openStudent.student.id ? { ...s, ...patch } : s));
            persistRoster({ ...roster, students: { ...roster.students, [gid]: list } });
            setOpenStudent({ ...openStudent, student: { ...openStudent.student, ...patch } });
          }}
        />
      )}
      </>)}
    </div>
  );
}

/* ============ MOVIMIENTOS (altas y bajas por mes/año) ============ */
function MovimientosView({ roster }) {
  const [year, setYear] = useState(new Date().getFullYear());
  const [openBaja, setOpenBaja] = useState(null); // {name, reason, date}

  const allStudents = useMemo(() => {
    const out = [];
    GROUPS.forEach((g) => (roster.students[g.id] || []).forEach((s) => out.push({ ...s, groupLabel: g.label })));
    return out;
  }, [roster.students]);

  const years = useMemo(() => {
    const set = new Set();
    allStudents.forEach((s) => { if (s.fechaAlta) set.add(parseISO(s.fechaAlta).getFullYear()); if (s.bajaDate) set.add(parseISO(s.bajaDate).getFullYear()); });
    set.add(new Date().getFullYear());
    return [...set].sort();
  }, [allStudents]);

  const altasPorMes = useMemo(() => {
    const m = {};
    allStudents.forEach((s) => {
      if (!s.fechaAlta) return;
      const d = parseISO(s.fechaAlta);
      if (d.getFullYear() !== year) return;
      const key = pad2(d.getMonth() + 1);
      (m[key] = m[key] || []).push(s);
    });
    return m;
  }, [allStudents, year]);

  const bajasPorMes = useMemo(() => {
    const m = {};
    allStudents.forEach((s) => {
      if (s.active !== false || !s.bajaDate) return;
      const d = parseISO(s.bajaDate);
      if (d.getFullYear() !== year) return;
      const key = pad2(d.getMonth() + 1);
      (m[key] = m[key] || []).push(s);
    });
    return m;
  }, [allStudents, year]);

  const totalAltas = Object.values(altasPorMes).reduce((a, arr) => a + arr.length, 0);
  const totalBajas = Object.values(bajasPorMes).reduce((a, arr) => a + arr.length, 0);

  return (
    <div>
      <Card style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontSize: 13, color: T.muted }}>{totalAltas} altas · {totalBajas} bajas en {year}</div>
        <div style={{ display: "flex", gap: 4 }}>
          {years.map((y) => (
            <button key={y} onClick={() => setYear(y)} style={{ padding: "5px 10px", borderRadius: 7, fontSize: 12.5, fontWeight: 600, border: `1px solid ${y === year ? T.ink : T.line}`, background: y === year ? T.ink : "#fff", color: y === year ? "#fff" : T.ink }}>{y}</button>
          ))}
        </div>
      </Card>

      <Card>
        <Label>Altas por mes</Label>
        {MONTH_NAMES.map((mn, i) => {
          const key = pad2(i + 1);
          const list = altasPorMes[key] || [];
          if (list.length === 0) return null;
          return (
            <div key={key} style={{ padding: "8px 0", borderTop: `1px solid ${T.paperDim}` }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 4 }}>{mn} — {list.length}</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {list.map((s) => (
                  <span key={s.id} style={{ background: T.greenBg, color: T.green, borderRadius: 999, padding: "3px 10px", fontSize: 12 }}>{s.name}</span>
                ))}
              </div>
            </div>
          );
        })}
        {totalAltas === 0 && <div style={{ fontSize: 13, color: T.muted }}>Sin altas registradas en {year}.</div>}
      </Card>

      <Card>
        <Label>Bajas por mes</Label>
        {MONTH_NAMES.map((mn, i) => {
          const key = pad2(i + 1);
          const list = bajasPorMes[key] || [];
          if (list.length === 0) return null;
          return (
            <div key={key} style={{ padding: "8px 0", borderTop: `1px solid ${T.paperDim}` }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, marginBottom: 4 }}>{mn} — {list.length}</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {list.map((s) => (
                  <button key={s.id} onClick={() => setOpenBaja(s)} style={{ background: T.redBg, color: T.red, borderRadius: 999, padding: "3px 10px", fontSize: 12, border: "none" }}>{s.name}</button>
                ))}
              </div>
            </div>
          );
        })}
        {totalBajas === 0 && <div style={{ fontSize: 13, color: T.muted }}>Sin bajas registradas en {year}.</div>}
      </Card>

      {openBaja && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60 }} onClick={() => setOpenBaja(null)}>
          <div style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 360, padding: 20 }} onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
              <div style={{ fontWeight: 800, fontSize: 15 }}>{openBaja.name}</div>
              <button onClick={() => setOpenBaja(null)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={18} /></button>
            </div>
            <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 10 }}>{openBaja.groupLabel} · baja el {fmtDateShort(openBaja.bajaDate)}</div>
            <div style={{ fontSize: 14 }}>{openBaja.bajaReason || "Sin motivo registrado."}</div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============ FICHA DE ALUMNO/A ============ */
function MiniMonthCalendar({ cursor, group, sessByDate, studentId }) {
  const cyear = cursor.getFullYear(), cmonth = cursor.getMonth();
  const firstDay = new Date(cyear, cmonth, 1);
  const startOffset = (firstDay.getDay() + 6) % 7;
  const daysInMonth = new Date(cyear, cmonth + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontWeight: 700, fontSize: 13, textTransform: "capitalize", marginBottom: 6 }}>{MONTH_NAMES[cmonth]} {cyear}</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 3 }}>
        {DOW_NAMES.map((d) => <div key={d} style={{ textAlign: "center", fontSize: 9.5, fontWeight: 700, color: T.muted }}>{d}</div>)}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3 }}>
        {cells.map((d, i) => {
          if (d === null) return <div key={i} />;
          const dateStr = `${cyear}-${pad2(cmonth + 1)}-${pad2(d)}`;
          const dow = new Date(cyear, cmonth, d).getDay();
          const isClassDay = dow === group.dow;
          const s = sessByDate[dateStr];
          const v = s?.attendance?.[studentId];
          let bg = "transparent", color = T.ink, border = T.paperDim;
          if (isClassDay) {
            if (v === true) { bg = T.green; color = "#fff"; border = T.green; }
            else if (v === false) { bg = T.redBg; color = T.red; border = T.red; }
            else { border = T.accent; }
          }
          return <div key={i} style={{ aspectRatio: "1", borderRadius: 6, border: `1px solid ${border}`, background: bg, color, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: isClassDay ? 700 : 400 }}>{d}</div>;
        })}
      </div>
    </div>
  );
}

const PAGO_IMPORTES = [15, 25, 50];

/* ============ PAGOS (por grupo, un mes a la vez) ============ */
function PagosTab({ ctx }) {
  const { roster, payments, persistPayments } = ctx;
  // topKey: la clave de un día de DAYS ("lunes"...), o "extra" (clases sueltas), o "total".
  const [topKey, setTopKey] = useState(DAYS[0].key);
  const [groupId, setGroupId] = useState(GROUPS.find((g) => g.day === DAYS[0].key).id);
  const [sporadicId, setSporadicId] = useState(null);
  const [ym, setYm] = useState(yyyymm(new Date()));

  const clasesSueltas = roster.clasesSueltas || [];
  const isTotal = topKey === "total";
  const isExtra = topKey === "extra";
  const activeGroupId = isTotal ? null : isExtra ? sporadicId : groupId;
  const group = activeGroupId ? findGroup(roster, activeGroupId) : null;
  const students = group ? (roster.students[activeGroupId] || []).filter((s) => s.active !== false) : [];
  const admins = roster.teachers.filter((t) => t.isAdmin);

  // Mes actual a la izquierda, retrocediendo hacia la derecha — igual que en Asistencia.
  const months = useMemo(() => {
    const now = new Date();
    const arr = [];
    for (let i = 0; i < 12; i++) arr.push(yyyymm(new Date(now.getFullYear(), now.getMonth() - i, 1)));
    return arr;
  }, []);

  function setPayment(studentId, patch) {
    const cur = (payments.students[studentId] || {})[ym] || {};
    const next = { ...payments, students: { ...payments.students, [studentId]: { ...(payments.students[studentId] || {}), [ym]: { ...cur, ...patch } } } };
    persistPayments(next);
  }

  const rows = students.map((s) => ({ student: s, p: (payments.students[s.id] || {})[ym] || {} }));
  const pendientes = rows.filter((r) => !r.p.paid);

  const todosLosGrupos = [...GROUPS, ...clasesSueltas];
  function grupoCompletoEnMes(g, m) {
    const sts = (roster.students[g.id] || []).filter((s) => s.active !== false);
    if (sts.length === 0) return false;
    return sts.every((s) => (payments.students[s.id] || {})[m]?.paid);
  }
  function montoGrupoEnMes(g, m) {
    let sum = 0;
    (roster.students[g.id] || []).filter((s) => s.active !== false).forEach((s) => {
      const p = (payments.students[s.id] || {})[m] || {};
      if (p.paid) sum += Number(p.amount || 0);
    });
    return sum;
  }
  function totalPorMes(m) {
    return todosLosGrupos.reduce((sum, g) => sum + montoGrupoEnMes(g, m), 0);
  }
  // Cuánto ha cobrado cada administrador ese mes (según a quién se marcó el pago), para
  // desglosarlo junto al total en la vista TOTAL.
  function montoPorProfesorEnMes(m) {
    const porProfesor = {};
    todosLosGrupos.forEach((g) => {
      (roster.students[g.id] || []).filter((s) => s.active !== false).forEach((s) => {
        const p = (payments.students[s.id] || {})[m] || {};
        if (p.paid && p.amount) {
          const quien = p.paidTo || "Sin especificar";
          porProfesor[quien] = (porProfesor[quien] || 0) + Number(p.amount);
        }
      });
    });
    return Object.entries(porProfesor).sort((a, b) => b[1] - a[1]);
  }
  // Para pintar en verde, en la tira de meses, el mes ya cobrado del todo del grupo elegido.
  const mesesCompletosDelGrupo = useMemo(() => {
    if (!group || students.length === 0) return new Set();
    const set = new Set();
    months.forEach((m) => { if (students.every((s) => (payments.students[s.id] || {})[m]?.paid)) set.add(m); });
    return set;
  }, [months, students, payments.students, group]);

  return (
    <div>
      <Card>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {DAYS.map((d) => (
            <button key={d.key} onClick={() => { setTopKey(d.key); const g = GROUPS.find((x) => x.day === d.key); if (g) setGroupId(g.id); }} style={{
              padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 600,
              border: `1px solid ${topKey === d.key ? T.ink : T.line}`,
              background: topKey === d.key ? T.ink : "#fff", color: topKey === d.key ? "#fff" : T.ink,
            }}>{d.label}</button>
          ))}
          <button onClick={() => setTopKey("extra")} style={{
            padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 600,
            border: `1px solid ${isExtra ? T.ink : T.line}`,
            background: isExtra ? T.ink : "#fff", color: isExtra ? "#fff" : T.ink,
          }}>Clases extra</button>
          <button onClick={() => setTopKey("total")} style={{
            padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 800,
            border: `1px solid ${isTotal ? T.accent : T.line}`,
            background: isTotal ? T.accent : "#fff", color: isTotal ? "#fff" : T.accent,
          }}>TOTAL</button>
        </div>

        {!isTotal && !isExtra && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 10 }}>
            {GROUPS.filter((g) => g.day === topKey).map((g) => (
              <button key={g.id} onClick={() => setGroupId(g.id)} style={{
                padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 600,
                border: `1px solid ${groupId === g.id ? T.ink : T.line}`,
                background: groupId === g.id ? T.ink : "#fff", color: groupId === g.id ? "#fff" : T.ink,
              }}>{g.level}</button>
            ))}
          </div>
        )}
        {isExtra && (
          clasesSueltas.length === 0 ? (
            <div style={{ fontSize: 13, color: T.muted, marginTop: 10 }}>Todavía no has creado ninguna clase suelta. Se crean desde el Calendario, al clicar un día → "Crear clase nueva".</div>
          ) : (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 10 }}>
              {clasesSueltas.map((g) => (
                <button key={g.id} onClick={() => setSporadicId(g.id)} style={{
                  padding: "7px 13px", borderRadius: 8, fontSize: 13, fontWeight: 600,
                  border: `1px solid ${sporadicId === g.id ? T.ink : T.line}`,
                  background: sporadicId === g.id ? T.ink : "#fff", color: sporadicId === g.id ? "#fff" : T.ink,
                }}>{g.label}</button>
              ))}
            </div>
          )
        )}
      </Card>

      <Card style={{ padding: "10px 0" }}>
        <div className="tab-scroll" style={{ padding: "0 14px", gap: 6 }}>
          {months.map((m) => {
            const selected = ym === m;
            const completo = !isTotal && mesesCompletosDelGrupo.has(m);
            const monto = !isTotal && group ? montoGrupoEnMes(group, m) : 0;
            return (
              <button key={m} onClick={() => setYm(m)} style={{
                padding: "8px 14px", borderRadius: 999, border: "none", fontSize: 12.5, fontWeight: 700, textTransform: "capitalize",
                background: selected ? T.ink : completo ? T.greenBg : T.paperDim,
                color: selected ? "#fff" : completo ? T.green : T.muted, flexShrink: 0, whiteSpace: "nowrap",
              }}>{monthLabel(m)}{completo ? ` · ${monto} €` : ""}</button>
            );
          })}
        </div>
      </Card>

      {isTotal ? (
        months.map((m) => {
          const porProfesor = montoPorProfesorEnMes(m);
          return (
          <Card key={m}>
            <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
              <Label>{monthLabel(m)}</Label>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 18, fontWeight: 800 }}>{totalPorMes(m)} €</div>
                {porProfesor.length > 0 && (
                  <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
                    {porProfesor.map(([nombre, monto]) => `${nombre}: ${monto} €`).join(" · ")}
                  </div>
                )}
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8 }}>
              {todosLosGrupos.map((g) => {
                const sts = (roster.students[g.id] || []).filter((s) => s.active !== false);
                const completo = grupoCompletoEnMes(g, m);
                const monto = montoGrupoEnMes(g, m);
                return (
                  <div key={g.id} style={{
                    padding: "10px 12px", borderRadius: 10, fontSize: 12.5, fontWeight: 700,
                    border: `1px solid ${completo ? T.green : T.line}`,
                    background: completo ? T.green : "#fff", color: completo ? "#fff" : T.ink,
                  }}>
                    {g.label}
                    {sts.length === 0 ? (
                      <div style={{ fontSize: 10.5, fontWeight: 500, opacity: 0.75, marginTop: 2 }}>Sin alumnos</div>
                    ) : (
                      <div style={{ fontSize: 11, fontWeight: 700, opacity: 0.9, marginTop: 2 }}>{monto} €</div>
                    )}
                  </div>
                );
              })}
            </div>
          </Card>
          );
        })
      ) : group && (
        <>
          <Card style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
            <div style={{ fontWeight: 800, fontSize: 15, textTransform: "capitalize" }}>{group.label} · {monthLabel(ym)}</div>
            {students.length > 0 && (
              pendientes.length === 0
                ? <Badge tone="green">Todos han pagado</Badge>
                : <Badge tone="accent">{pendientes.length} sin pagar</Badge>
            )}
          </Card>

          {students.length === 0 ? (
            <Card style={{ textAlign: "center", color: T.muted }}>Este grupo no tiene alumnos activos todavía.</Card>
          ) : (
            rows.map(({ student, p }) => (
              <Card key={student.id} style={{ padding: 14 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <Avatar name={student.name} photo={student.photo} size={32} />
                  <div style={{ flex: 1, fontSize: 14, fontWeight: 700 }}>{student.name}</div>
                  <button onClick={() => setPayment(student.id, { paid: !p.paid })} style={attendanceBoxStyle(p.paid, 28)}>
                    {p.paid ? <Check size={17} color="#fff" /> : null}
                  </button>
                </div>
                {p.paid && (
                  <div className="two-col" style={{ gap: 8, marginTop: 10 }}>
                    <select style={inputStyle()} value={p.amount || ""} onChange={(e) => setPayment(student.id, { amount: e.target.value })}>
                      <option value="">Cantidad</option>
                      {PAGO_IMPORTES.map((v) => <option key={v} value={v}>{v} €</option>)}
                    </select>
                    <select style={inputStyle()} value={p.paidTo || ""} onChange={(e) => setPayment(student.id, { paidTo: e.target.value })}>
                      <option value="">¿A quién?</option>
                      {admins.map((a) => <option key={a.name} value={a.name}>{a.name}</option>)}
                    </select>
                  </div>
                )}
              </Card>
            ))
          )}
        </>
      )}
    </div>
  );
}

function StudentModal({ student, group, ctx, onClose, onSaveDetails }) {
  const { roster, sessions, payments, persistPayments, isAdmin } = ctx;
  const [sub, setSub] = useState("asistencia");
  const [range, setRange] = useState(1); // 1 or 3 months
  const [cursor, setCursor] = useState(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  const [phone, setPhone] = useState(student.phone || "");
  const [notes, setNotes] = useState(student.notes || "");
  const [fechaAlta, setFechaAlta] = useState(student.fechaAlta || "");

  function antiguedadTexto(fecha) {
    if (!fecha) return null;
    const d = parseISO(fecha);
    const now = new Date();
    let months = (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth());
    if (now.getDate() < d.getDate()) months--;
    if (months < 0) return "fecha futura";
    if (months < 1) return "menos de un mes";
    const years = Math.floor(months / 12);
    const rem = months % 12;
    const parts = [];
    if (years > 0) parts.push(`${years} año${years !== 1 ? "s" : ""}`);
    if (rem > 0) parts.push(`${rem} mes${rem !== 1 ? "es" : ""}`);
    return parts.join(" y ");
  }

  const sessByDate = useMemo(() => {
    const m = {};
    sessions.filter((s) => s.groupId === group.id).forEach((s) => (m[s.date] = s));
    return m;
  }, [sessions, group.id]);

  let presentCount = 0, absentCount = 0;
  Object.values(sessByDate).forEach((s) => {
    const v = s.attendance?.[student.id];
    if (v === true) presentCount++; else if (v === false) absentCount++;
  });

  const months = range === 1 ? [cursor] : [cursor, new Date(cursor.getFullYear(), cursor.getMonth() - 1, 1), new Date(cursor.getFullYear(), cursor.getMonth() - 2, 1)];

  const studentPayments = payments.students[student.id] || {};
  function setPayment(ym, patch) {
    const next = { ...payments, students: { ...payments.students, [student.id]: { ...studentPayments, [ym]: { ...(studentPayments[ym] || {}), ...patch } } } };
    persistPayments(next);
  }
  const paymentMonths = useMemo(() => {
    const now = new Date();
    const arr = [];
    for (let i = 0; i <= 5; i++) arr.push(yyyymm(new Date(now.getFullYear(), now.getMonth() - i, 1)));
    return arr;
  }, []);

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 55 }} onClick={onClose}>
      <div className="modal-sheet-pad" style={{ background: "#fff", borderRadius: "16px 16px 0 0", width: "100%", maxWidth: 480, maxHeight: "88vh", overflowY: "auto", padding: 20 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <label style={{ position: "relative", cursor: isAdmin ? "pointer" : "default" }}>
              <Avatar name={student.name} photo={student.photo} size={52} />
              {isAdmin && (
                <span style={{ position: "absolute", bottom: -2, right: -2, background: T.ink, color: "#fff", borderRadius: "50%", width: 20, height: 20, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <Pencil size={11} />
                </span>
              )}
              {isAdmin && (
                <input type="file" accept="image/*" style={{ display: "none" }} onChange={async (e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  try { const dataUrl = await fileToResizedDataURL(file); onSaveDetails({ photo: dataUrl }); }
                  catch { /* ignora si falla la lectura de la imagen */ }
                  e.target.value = "";
                }} />
              )}
            </label>
            <div style={{ fontWeight: 800, fontSize: 16 }}>{student.name}</div>
          </div>
          <button onClick={onClose} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={20} /></button>
        </div>
        <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 10 }}>{group.label} · {presentCount} asistencias / {absentCount} faltas</div>

        {isAdmin && (
          <div className="two-col" style={{ gap: 8, marginBottom: 10 }}>
            <div>
              <Label>Teléfono</Label>
              <input style={inputStyle()} value={phone} onChange={(e) => setPhone(e.target.value)} onBlur={() => onSaveDetails({ phone })} placeholder="—" />
            </div>
            <div>
              <Label>Comentario</Label>
              <input style={inputStyle()} value={notes} onChange={(e) => setNotes(e.target.value)} onBlur={() => onSaveDetails({ notes })} placeholder="—" />
            </div>
          </div>
        )}
        {isAdmin && (
          <div style={{ marginBottom: 14 }}>
            <Label>En la escuela desde</Label>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <input type="date" style={inputStyle({ maxWidth: 170 })} value={fechaAlta} onChange={(e) => { setFechaAlta(e.target.value); onSaveDetails({ fechaAlta: e.target.value }); }} />
              {antiguedadTexto(fechaAlta) && <span style={{ fontSize: 12.5, color: T.muted }}>({antiguedadTexto(fechaAlta)})</span>}
            </div>
          </div>
        )}
        {!isAdmin && (student.phone || student.notes) && (
          <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 10 }}>
            {student.phone && <div>Tel: {student.phone}</div>}
            {student.notes && <div>{student.notes}</div>}
          </div>
        )}

        <SubTabs tabs={[{ id: "asistencia", label: "Asistencia" }, { id: "pagos", label: "Pagos" }]} active={sub} setActive={setSub} />

        {sub === "asistencia" && (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
              <button style={btn("ghost", { padding: "6px 10px" })} onClick={() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() - 1, 1))}><ChevronLeft size={15} /></button>
              <div style={{ display: "flex", gap: 4 }}>
                <button style={btn(range === 1 ? "primary" : "ghost", { padding: "5px 10px", fontSize: 11.5 })} onClick={() => setRange(1)}>1 mes</button>
                <button style={btn(range === 3 ? "primary" : "ghost", { padding: "5px 10px", fontSize: 11.5 })} onClick={() => setRange(3)}>3 meses</button>
              </div>
              <button style={btn("ghost", { padding: "6px 10px" })} onClick={() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1))}><ChevronRight size={15} /></button>
            </div>
            {months.map((m, i) => <MiniMonthCalendar key={i} cursor={m} group={group} sessByDate={sessByDate} studentId={student.id} />)}
            <div style={{ display: "flex", gap: 14, fontSize: 11, color: T.muted, flexWrap: "wrap" }}>
              <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 11, height: 11, borderRadius: 3, background: T.green }} /> Asistió</span>
              <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 11, height: 11, borderRadius: 3, background: T.redBg, border: `1px solid ${T.red}` }} /> Faltó</span>
              <span style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ width: 11, height: 11, borderRadius: 3, border: `1px solid ${T.accent}` }} /> Sin registrar</span>
            </div>
          </div>
        )}

        {sub === "pagos" && (
          <div>
            {paymentMonths.map((ym) => {
              const p = studentPayments[ym] || {};
              return (
                <div key={ym} style={{ border: `1px solid ${T.paperDim}`, borderRadius: 10, padding: 10, marginBottom: 8 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, textTransform: "capitalize" }}>{monthLabel(ym)}</div>
                    {isAdmin ? (
                      <button onClick={() => setPayment(ym, { paid: !p.paid })} style={attendanceBoxStyle(p.paid, 24)}>
                        {p.paid ? <Check size={14} color="#fff" /> : null}
                      </button>
                    ) : (
                      <Badge tone={p.paid ? "green" : "neutral"}>{p.paid ? "Pagado" : "Pendiente"}</Badge>
                    )}
                  </div>
                  {isAdmin && p.paid && (
                    <div className="two-col" style={{ gap: 8 }}>
                      <select style={inputStyle({ padding: "6px 9px", fontSize: 12.5 })} value={p.amount || ""} onChange={(e) => setPayment(ym, { amount: e.target.value })}>
                        <option value="">Cantidad</option>
                        {PAGO_IMPORTES.map((v) => <option key={v} value={v}>{v} €</option>)}
                      </select>
                      <select style={inputStyle({ padding: "6px 9px", fontSize: 12.5 })} value={p.paidTo || ""} onChange={(e) => setPayment(ym, { paidTo: e.target.value })}>
                        <option value="">¿A quién?</option>
                        {roster.teachers.filter((t) => t.isAdmin).map((t) => <option key={t.name} value={t.name}>{t.name}</option>)}
                      </select>
                    </div>
                  )}
                  {!isAdmin && p.paid && (p.amount || p.paidTo) && (
                    <div style={{ fontSize: 12, color: T.muted }}>{p.amount ? `${p.amount} € · ` : ""}{p.paidTo}</div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

/* ============ CLASES (listado con filtros) ============ */
function ClasesTab({ ctx }) {
  const { sessions, persistSessions, roster, askConfirm } = ctx;
  const [fDay, setFDay] = useState("todos");
  const [fLevel, setFLevel] = useState("todos");
  const [fEstado, setFEstado] = useState("todos");
  const [fProfesor, setFProfesor] = useState("todos");
  const [openSession, setOpenSession] = useState(null);

  const rows = sessions.filter((s) => {
    const g = findGroup(roster, s.groupId);
    if (fDay !== "todos" && (!g || g.day !== fDay)) return false;
    if (fLevel !== "todos" && s.groupId !== fLevel) return false;
    const given = isGiven(s);
    if (fEstado === "impartidas" && !given) return false;
    if (fEstado === "programadas" && given) return false;
    if (fProfesor !== "todos" && s.profesor !== fProfesor) return false;
    return true;
  }).slice().sort((a, b) => b.date.localeCompare(a.date));

  function removeSession(id) {
    askConfirm("¿Eliminar esta clase? No se puede deshacer.", async () => {
      await persistSessions(sessions.filter((s) => s.id !== id));
    });
  }

  return (
    <div>
      <Card>
        <div className="two-col" style={{ gap: 8 }}>
          <select style={inputStyle()} value={fDay} onChange={(e) => setFDay(e.target.value)}>
            <option value="todos">Todos los días</option>
            {DAYS.map((d) => <option key={d.key} value={d.key}>{d.label}</option>)}
          </select>
          <select style={inputStyle()} value={fLevel} onChange={(e) => setFLevel(e.target.value)}>
            <option value="todos">Todos los grupos</option>
            {GROUPS.map((g) => <option key={g.id} value={g.id}>{g.label}</option>)}
            {(roster.clasesSueltas || []).length > 0 && <option disabled>── Clases sueltas ──</option>}
            {(roster.clasesSueltas || []).map((g) => <option key={g.id} value={g.id}>{g.label}</option>)}
          </select>
          <select style={inputStyle()} value={fEstado} onChange={(e) => setFEstado(e.target.value)}>
            <option value="todos">Todos los estados</option>
            <option value="impartidas">Impartidas</option>
            <option value="programadas">Programadas</option>
          </select>
          <select style={inputStyle()} value={fProfesor} onChange={(e) => setFProfesor(e.target.value)}>
            <option value="todos">Todos los profesores</option>
            {roster.teachers.map((t) => <option key={t.name} value={t.name}>{t.name}</option>)}
          </select>
        </div>
      </Card>

      {rows.length === 0 && <Card style={{ textAlign: "center", color: T.muted }}>No hay clases con estos filtros.</Card>}

      {rows.map((s) => {
        const g = findGroup(roster, s.groupId);
        const given = isGiven(s);
        return (
          <Card key={s.id} style={{ cursor: "pointer" }} >
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }} onClick={() => setOpenSession({ date: s.date, groupId: s.groupId })}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <span style={{ fontSize: 10.5, fontWeight: 800, padding: "3px 8px", borderRadius: 999, background: given ? T.greenBg : T.accentDim, color: given ? T.green : T.accent }}>{given ? "IMPARTIDA" : "PROGRAMADA"}</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14.5, textTransform: "capitalize" }}>{fmtDateLong(s.date)}</div>
                  <div style={{ fontSize: 12.5, color: T.muted }}>{g?.label || "Grupo eliminado"}</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 16, alignItems: "center", fontSize: 13 }}>
                <div><span style={{ color: T.muted }}>Profesor/a:</span> <strong>{s.profesor || "—"}</strong></div>
                <div><span style={{ color: T.muted }}>Tipo:</span> <strong>{s.tipoClase || "—"}</strong></div>
                <button onClick={(e) => { e.stopPropagation(); removeSession(s.id); }} style={btn("danger", { padding: "6px 9px" })}><Trash2 size={13} /></button>
              </div>
            </div>
            {s.observaciones && <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${T.paperDim}`, fontSize: 13 }}><span style={{ color: T.muted }}>Observaciones: </span>{s.observaciones}</div>}
          </Card>
        );
      })}

      {openSession && <ClassEditor date={openSession.date} groupId={openSession.groupId} ctx={ctx} onClose={() => setOpenSession(null)} />}
    </div>
  );
}

/* ============ PROFESORES ============ */
function ProfesoresTab({ ctx }) {
  const { roster, sessions, payments, persistPayments, addTeacher, updateTeacher, removeTeacher, askConfirm, resetSessions } = ctx;
  const [newTeacher, setNewTeacher] = useState("");
  const [newColor, setNewColor] = useState(TEACHER_PALETTE[0]);
  const [newPassword, setNewPassword] = useState("");
  const [year, setYear] = useState(new Date().getFullYear());
  const [openTeacherName, setOpenTeacherName] = useState(null);
  const openTeacher = roster.teachers.find((t) => t.name === openTeacherName) || null;

  const years = useMemo(() => {
    const set = new Set(sessions.map((s) => parseISO(s.date).getFullYear()));
    set.add(new Date().getFullYear());
    return [...set].sort();
  }, [sessions]);

  const todayStr = iso(new Date());
  const counts = useMemo(() => {
    const m = {};
    roster.teachers.forEach((t) => (m[t.name] = { months: Array(12).fill(0), future: 0 }));
    sessions.forEach((s) => {
      if (!s.profesor) return;
      if (!m[s.profesor]) m[s.profesor] = { months: Array(12).fill(0), future: 0 };
      const given = isGiven(s);
      if (s.date > todayStr && !given) m[s.profesor].future++;
      else { const d = parseISO(s.date); if (d.getFullYear() === year) m[s.profesor].months[d.getMonth()]++; }
    });
    return m;
  }, [sessions, roster.teachers, year, todayStr]);

  const paymentMonths = useMemo(() => {
    const now = new Date();
    const arr = [];
    for (let i = 0; i <= 5; i++) arr.push(yyyymm(new Date(now.getFullYear(), now.getMonth() - i, 1)));
    return arr;
  }, []);
  function setTeacherPayment(name, ym, patch) {
    const cur = payments.teachers[name] || {};
    persistPayments({ ...payments, teachers: { ...payments.teachers, [name]: { ...cur, [ym]: { ...(cur[ym] || {}), ...patch } } } });
  }

  return (
    <div>
      <Card>
        <Label>Profesores</Label>
        <p style={{ fontSize: 12.5, color: T.muted, marginTop: -4, marginBottom: 12 }}>Añade aquí a cada profesor/a con un color propio para identificarlo en el calendario.</p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
          {roster.teachers.map((t) => (
            <button key={t.name} onClick={() => setOpenTeacherName(t.name)} style={{ display: "inline-flex", alignItems: "center", gap: 6, background: T.paperDim, borderRadius: 999, padding: "5px 7px", fontSize: 13.5, border: "none" }}>
              <Avatar name={t.name} photo={t.photo} size={24} color={t.color} />
              {t.name}
              {t.isAdmin && <span style={{ fontSize: 9, fontWeight: 800, color: T.accent }}>ADMIN</span>}
              <span onClick={(e) => { e.stopPropagation(); removeTeacher(t.name); }} style={{ color: T.muted, display: "flex", padding: 3 }}><X size={13} /></span>
            </button>
          ))}
          {roster.teachers.length === 0 && <span style={{ fontSize: 13, color: T.muted }}>Todavía no hay ningún profesor/a añadido.</span>}
        </div>
        <div className="two-col" style={{ gap: 8, marginBottom: 10 }}>
          <input style={inputStyle()} placeholder="Nombre del profesor/a" value={newTeacher} onChange={(e) => setNewTeacher(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { addTeacher(newTeacher, newColor, newPassword); setNewTeacher(""); setNewPassword(""); } }} />
          <input style={inputStyle()} placeholder="Contraseña (opcional, si no se genera sola)" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { addTeacher(newTeacher, newColor, newPassword); setNewTeacher(""); setNewPassword(""); } }} />
        </div>
        <button style={btn("primary", { marginBottom: 10 })} onClick={() => { addTeacher(newTeacher, newColor, newPassword); setNewTeacher(""); setNewPassword(""); }}><Plus size={15} /> Añadir</button>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {TEACHER_PALETTE.map((c) => (
            <button key={c} onClick={() => setNewColor(c)} style={{ width: 24, height: 24, borderRadius: "50%", background: c, border: newColor === c ? `2px solid ${T.ink}` : "2px solid transparent" }} />
          ))}
        </div>
      </Card>

      <Card style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: 16 }}>
          <Label>Clases impartidas por mes</Label>
          <div style={{ display: "flex", gap: 4 }}>
            {years.map((y) => (
              <button key={y} onClick={() => setYear(y)} style={{ padding: "5px 10px", borderRadius: 7, fontSize: 12.5, fontWeight: 600, border: `1px solid ${y === year ? T.ink : T.line}`, background: y === year ? T.ink : "#fff", color: y === year ? "#fff" : T.ink }}>{y}</button>
            ))}
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", fontSize: 12.5 }}>
            <thead>
              <tr style={{ background: T.paperDim }}>
                <th style={{ textAlign: "left", padding: "8px 14px" }}>Profesor/a</th>
                {MONTH_NAMES.map((m) => <th key={m} style={{ padding: "8px 4px", fontWeight: 600, minWidth: 42 }}>{m.slice(0, 3)}</th>)}
                <th style={{ padding: "8px 10px", fontWeight: 800 }}>Total {year}</th>
                <th style={{ padding: "8px 10px", fontWeight: 800, color: T.accent }}>Programadas</th>
              </tr>
            </thead>
            <tbody>
              {Object.keys(counts).length === 0 && (<tr><td colSpan={15} style={{ padding: 16, textAlign: "center", color: T.muted }}>Sin datos todavía.</td></tr>)}
              {Object.entries(counts).map(([name, data]) => (
                <tr key={name} style={{ borderTop: `1px solid ${T.paperDim}` }}>
                  <td style={{ padding: "8px 14px", fontWeight: 600 }}>{name}</td>
                  {data.months.map((n, i) => <td key={i} style={{ textAlign: "center", padding: "8px 4px", color: n ? T.ink : T.muted }}>{n || "—"}</td>)}
                  <td style={{ textAlign: "center", padding: "8px 10px", fontWeight: 800 }}>{data.months.reduce((a, b) => a + b, 0)}</td>
                  <td style={{ textAlign: "center", padding: "8px 10px", fontWeight: 800, color: data.future ? T.accent : T.muted }}>{data.future || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      {openTeacher && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,10,0.45)", display: "flex", alignItems: "flex-end", justifyContent: "center", zIndex: 55 }} onClick={() => setOpenTeacherName(null)}>
          <div style={{ background: "#fff", borderRadius: "16px 16px 0 0", width: "100%", maxWidth: 420, maxHeight: "85vh", overflowY: "auto", padding: 20 }} onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <label style={{ position: "relative", cursor: "pointer" }}>
                  <Avatar name={openTeacher.name} photo={openTeacher.photo} size={52} color={openTeacher.color} />
                  <span style={{ position: "absolute", bottom: -2, right: -2, background: T.ink, color: "#fff", borderRadius: "50%", width: 20, height: 20, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Pencil size={11} />
                  </span>
                  <input type="file" accept="image/*" style={{ display: "none" }} onChange={async (e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    try { const dataUrl = await fileToResizedDataURL(file); updateTeacher(openTeacher.name, { photo: dataUrl }); }
                    catch { /* ignora si falla la lectura de la imagen */ }
                    e.target.value = "";
                  }} />
                </label>
                <div style={{ fontWeight: 800, fontSize: 16, display: "flex", alignItems: "center", gap: 8 }}><span style={{ width: 12, height: 12, borderRadius: "50%", background: openTeacher.color }} />{openTeacher.name}</div>
              </div>
              <button onClick={() => setOpenTeacherName(null)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><X size={20} /></button>
            </div>

            <Label>Color</Label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
              {TEACHER_PALETTE.map((c) => (
                <button key={c} onClick={() => updateTeacher(openTeacher.name, { color: c })} style={{ width: 26, height: 26, borderRadius: "50%", background: c, border: openTeacher.color === c ? `2px solid ${T.ink}` : "2px solid transparent" }} />
              ))}
            </div>

            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13.5, marginBottom: 16 }}>
              <input type="checkbox" checked={!!openTeacher.isAdmin} onChange={(e) => updateTeacher(openTeacher.name, { isAdmin: e.target.checked })} />
              Hacer administrador/a (verá y podrá editar todas las pestañas)
            </label>

            <Label>Contraseña personal</Label>
            <p style={{ fontSize: 11.5, color: T.muted, marginTop: -4, marginBottom: 8 }}>Escribe la que quieras — es la que {openTeacher.name} usará para entrar.</p>
            <PasswordField value={openTeacher.password} onSave={(v) => updateTeacher(openTeacher.name, { password: v })} />

            <Label>Pagos mensuales</Label>
            {paymentMonths.map((ym) => {
              const p = (payments.teachers[openTeacher.name] || {})[ym] || {};
              return (
                <div key={ym} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 0", borderTop: `1px solid ${T.paperDim}` }}>
                  <span style={{ fontSize: 13, textTransform: "capitalize" }}>{monthLabel(ym)}</span>
                  <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12 }}>
                    <input type="checkbox" checked={!!p.paid} onChange={(e) => setTeacherPayment(openTeacher.name, ym, { paid: e.target.checked })} /> Pagado
                  </label>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

/* ============ DINÁMICAS (catálogo maestro) ============ */
function broadLevel(level) { return (level || "").replace(/\s*\d+\s*$/, "").trim(); }

function DinamicasTab({ ctx }) {
  const { roster, sessions, addDinamica, removeDinamica, saveTipoClase, removeTipoClase, askConfirm } = ctx;
  const [newText, setNewText] = useState("");
  const [mode, setMode] = useState("catalogo"); // catalogo | porgrupo | pornivel | portipo | tipos
  const [vista, setVista] = useState("lista"); // lista | matriz (solo aplica en modo catalogo)
  const [search, setSearch] = useState("");
  const [editingTipo, setEditingTipo] = useState(null); // numero en edición, o "new"
  const [tipoDraft, setTipoDraft] = useState(null);
  const [newDinTipo, setNewDinTipo] = useState("");

  const doneStats = useMemo(() => {
    const m = {}; // dinamicaId -> {doneCount, groups:Set}
    sessions.forEach((s) => {
      (s.dinamicas || []).forEach((d) => {
        if (!m[d.id]) m[d.id] = { doneCount: 0, groups: new Set() };
        if (d.done) { m[d.id].doneCount++; m[d.id].groups.add(s.groupId); }
      });
    });
    return m;
  }, [sessions]);

  // veces que se ha hecho cada dinámica, dentro de sesiones de un GRUPO concreto (día+nivel)
  const doneByGroupAndDinamica = useMemo(() => {
    const m = {}; // `${groupId}|${dinamicaId}` -> count
    sessions.forEach((s) => {
      (s.dinamicas || []).forEach((d) => {
        if (!d.done) return;
        const key = `${s.groupId}|${d.id}`;
        m[key] = (m[key] || 0) + 1;
      });
    });
    return m;
  }, [sessions]);

  // veces que se ha hecho cada dinámica, dentro de sesiones de un NIVEL general (agrupando por ej. Avanzado 1 + Avanzado 2)
  const doneByLevelAndDinamica = useMemo(() => {
    const m = {}; // `${nivel}|${dinamicaId}` -> count
    sessions.forEach((s) => {
      const g = findGroup(roster, s.groupId);
      if (!g) return;
      const nivel = broadLevel(g.level);
      (s.dinamicas || []).forEach((d) => {
        if (!d.done) return;
        const key = `${nivel}|${d.id}`;
        m[key] = (m[key] || 0) + 1;
      });
    });
    return m;
  }, [sessions]);

  // días en los que se ha marcado como hecha alguna dinámica, y qué dinámicas se hicieron cada día
  // (para la vista de matriz dinámica × día)
  const doneDates = useMemo(() => {
    const set = new Set();
    sessions.forEach((s) => { if ((s.dinamicas || []).some((d) => d.done)) set.add(s.date); });
    return [...set].sort();
  }, [sessions]);
  const doneOnDate = useMemo(() => {
    const m = {}; // `${dinamicaId}|${date}` -> true
    sessions.forEach((s) => {
      (s.dinamicas || []).forEach((d) => { if (d.done) m[`${d.id}|${s.date}`] = true; });
    });
    return m;
  }, [sessions]);

  // veces que se ha hecho una dinámica concreta, DENTRO de sesiones de un tipo de clase concreto
  const doneByTipoAndDinamica = useMemo(() => {
    const m = {}; // `${numero}|${dinamicaId}` -> count
    sessions.forEach((s) => {
      if (!s.tipoClase) return;
      (s.dinamicas || []).forEach((d) => {
        if (!d.done) return;
        const key = `${s.tipoClase}|${d.id}`;
        m[key] = (m[key] || 0) + 1;
      });
    });
    return m;
  }, [sessions]);

  const niveles = useMemo(() => {
    const seen = [];
    GROUPS.forEach((g) => { const n = broadLevel(g.level); if (!seen.includes(n)) seen.push(n); });
    return seen;
  }, []);

  const sorted = roster.dinamicasCatalog
    .slice()
    .filter((d) => d.text.toLowerCase().includes(search.trim().toLowerCase()))
    .sort((a, b) => (doneStats[b.id]?.doneCount || 0) - (doneStats[a.id]?.doneCount || 0));

  function exportDinamicas() {
    const wb = XLSX.utils.book_new();

    // Hoja 1: dinámica (filas) x día en que se hizo (columnas), igual que la vista de matriz.
    const header1 = ["Dinámica", "Veces hecha", ...doneDates.map(fmtDateShort)];
    const rows1 = roster.dinamicasCatalog.map((d) => {
      const row = [d.text, doneStats[d.id]?.doneCount || 0];
      doneDates.forEach((dt) => row.push(doneOnDate[`${d.id}|${dt}`] ? "X" : ""));
      return row;
    });
    const ws1 = XLSX.utils.aoa_to_sheet([header1, ...rows1]);
    XLSX.utils.book_append_sheet(wb, ws1, "Dinámicas x día");

    // Hoja 2: qué dinámicas lleva cada tipo de clase.
    const header2 = ["Tipo de clase", "Dinámica"];
    const rows2 = [];
    roster.tiposClase.slice().sort((a, b) => a.numero - b.numero).forEach((tipo) => {
      (tipo.dinamicaIds || []).forEach((id) => {
        const cat = roster.dinamicasCatalog.find((c) => c.id === id);
        rows2.push([tipo.numero, cat ? cat.text : "(eliminada)"]);
      });
    });
    const ws2 = XLSX.utils.aoa_to_sheet([header2, ...rows2]);
    XLSX.utils.book_append_sheet(wb, ws2, "Por tipo de clase");

    XLSX.writeFile(wb, "Dinamicas.xlsx");
  }

  function startEditTipo(tipo) {
    setEditingTipo(tipo ? tipo.numero : "new");
    setTipoDraft(tipo ? { ...tipo, nivel: tipo.nivel || "Iniciación 1", dinamicaIds: [...tipo.dinamicaIds] } : { numero: "", nombre: "", nivel: "Iniciación 1", dinamicaIds: [], escenas: "" });
  }
  function addDinToTipoDraft() {
    const text = newDinTipo.trim();
    if (!text) return;
    const existingCat = roster.dinamicasCatalog.find((c) => c.text.toLowerCase() === text.toLowerCase());
    const item = existingCat || addDinamica(text);
    if (!item || tipoDraft.dinamicaIds.includes(item.id)) { setNewDinTipo(""); return; }
    setTipoDraft({ ...tipoDraft, dinamicaIds: [...tipoDraft.dinamicaIds, item.id] });
    setNewDinTipo("");
  }
  function removeDinFromTipoDraft(id) {
    setTipoDraft({ ...tipoDraft, dinamicaIds: tipoDraft.dinamicaIds.filter((x) => x !== id) });
  }
  function saveTipoDraft() {
    const numero = parseInt(tipoDraft.numero, 10);
    if (!numero || numero < 1 || numero > 40) return;
    if (editingTipo === "new" && roster.tiposClase.some((t) => t.numero === numero)) { alert("Ya existe un tipo de clase con ese número."); return; }
    saveTipoClase({ numero, nombre: "", nivel: tipoDraft.nivel || "Iniciación 1", dinamicaIds: tipoDraft.dinamicaIds, escenas: tipoDraft.escenas });
    setEditingTipo(null); setTipoDraft(null);
  }

  return (
    <div>
      <Card>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
          <button style={btn(mode === "catalogo" ? "primary" : "ghost", { padding: "7px 12px", fontSize: 12.5 })} onClick={() => setMode("catalogo")}>Todas las dinámicas</button>
          <button style={btn(mode === "porgrupo" ? "primary" : "ghost", { padding: "7px 12px", fontSize: 12.5 })} onClick={() => setMode("porgrupo")}>Por grupo</button>
          <button style={btn(mode === "pornivel" ? "primary" : "ghost", { padding: "7px 12px", fontSize: 12.5 })} onClick={() => setMode("pornivel")}>Por nivel</button>
          <button style={btn(mode === "portipo" ? "primary" : "ghost", { padding: "7px 12px", fontSize: 12.5 })} onClick={() => setMode("portipo")}>Por tipo de clase</button>
          <button style={btn(mode === "tipos" ? "primary" : "ghost", { padding: "7px 12px", fontSize: 12.5 })} onClick={() => setMode("tipos")}>Gestionar tipos de clase</button>
        </div>
        <button style={btn("ghost", { fontSize: 12.5 })} onClick={exportDinamicas}><List size={14} /> Descargar Excel de Dinámicas</button>
      </Card>

      {mode === "catalogo" && (
        <div>
          <Card>
            <Label>Añadir dinámica al catálogo</Label>
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <input style={inputStyle()} placeholder="Nombre de la dinámica" value={newText} onChange={(e) => setNewText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { addDinamica(newText); setNewText(""); } }} />
              <button style={btn("primary")} onClick={() => { addDinamica(newText); setNewText(""); }}><Plus size={15} /> Añadir</button>
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <div style={{ position: "relative", flex: 1 }}>
                <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: T.muted }} />
                <input style={inputStyle({ paddingLeft: 30 })} placeholder="Buscar dinámica…" value={search} onChange={(e) => setSearch(e.target.value)} />
              </div>
              <div style={{ display: "flex", gap: 4 }}>
                <button style={btn(vista === "lista" ? "primary" : "ghost", { padding: "8px 11px", fontSize: 12 })} onClick={() => setVista("lista")}>Lista</button>
                <button style={btn(vista === "matriz" ? "primary" : "ghost", { padding: "8px 11px", fontSize: 12 })} onClick={() => setVista("matriz")}>Matriz por día</button>
              </div>
            </div>
          </Card>

          {sorted.length === 0 && <Card style={{ textAlign: "center", color: T.muted }}>Sin dinámicas que coincidan.</Card>}

          {vista === "lista" && sorted.length > 0 && (
            <Card style={{ padding: 0, overflow: "hidden" }}>
              {sorted.map((d, i) => {
                const st = doneStats[d.id] || { doneCount: 0, groups: new Set() };
                const groupLabels = [...st.groups].map((gid) => GROUPS.find((g) => g.id === gid)?.level).filter(Boolean);
                return (
                  <div key={d.id} style={{
                    display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10,
                    padding: "8px 14px", borderTop: i === 0 ? "none" : `1px solid ${T.paperDim}`,
                  }}>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 13 }}>{d.text}</div>
                      <div style={{ fontSize: 11, color: T.muted }}>
                        {st.doneCount} {st.doneCount === 1 ? "vez" : "veces"}
                        {groupLabels.length > 0 && <> · {groupLabels.join(", ")}</>}
                      </div>
                    </div>
                    <button onClick={() => askConfirm("¿Eliminar esta dinámica del catálogo?", () => removeDinamica(d.id))} style={{ border: "none", background: "transparent", color: T.muted, display: "flex", padding: 4, flexShrink: 0 }}><Trash2 size={13} /></button>
                  </div>
                );
              })}
            </Card>
          )}

          {vista === "matriz" && sorted.length > 0 && (
            doneDates.length === 0 ? (
              <Card style={{ textAlign: "center", color: T.muted }}>Todavía no hay ninguna dinámica marcada como hecha en ninguna clase.</Card>
            ) : (
              <Card style={{ padding: 0, overflow: "hidden" }}>
                <div style={{ overflow: "auto", maxHeight: "70vh" }}>
                  <table style={{ fontSize: 12.5 }}>
                    <thead>
                      <tr>
                        <th style={{ position: "sticky", top: 0, left: 0, zIndex: 3, background: T.ink, color: "#fff", padding: "8px 12px", textAlign: "left", minWidth: 160, borderRight: `1px solid ${T.ink}` }}>Dinámica</th>
                        {doneDates.map((dt) => (
                          <th key={dt} style={{ position: "sticky", top: 0, zIndex: 2, background: T.ink, color: "#fff", padding: "6px 5px", fontWeight: 600, minWidth: 40, borderLeft: "1px solid rgba(255,255,255,0.15)" }}>{fmtDateShort(dt)}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {sorted.map((d, i) => (
                        <tr key={d.id} style={{ background: i % 2 ? T.paperDim : "#fff" }}>
                          <td style={{ position: "sticky", left: 0, zIndex: 1, background: i % 2 ? T.paperDim : "#fff", padding: "6px 12px", borderRight: `1px solid ${T.line}`, borderBottom: `1px solid ${T.paperDim}` }}>{d.text}</td>
                          {doneDates.map((dt) => {
                            const done = doneOnDate[`${d.id}|${dt}`];
                            return (
                              <td key={dt} style={{ padding: "6px 5px", textAlign: "center", borderBottom: `1px solid ${T.paperDim}`, borderLeft: `1px solid ${T.paperDim}`, background: done ? T.greenBg : "transparent" }}>
                                {done && <Check size={13} color={T.green} style={{ margin: "0 auto", display: "block" }} />}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            )
          )}
        </div>
      )}

      {mode === "porgrupo" && (
        <div>
          {GROUPS.map((g) => {
            const rows = roster.dinamicasCatalog
              .map((d) => ({ d, count: doneByGroupAndDinamica[`${g.id}|${d.id}`] || 0 }))
              .filter((r) => r.count > 0)
              .sort((a, b) => b.count - a.count);
            if (rows.length === 0) return null;
            return (
              <Card key={g.id}>
                <div style={{ fontWeight: 800, fontSize: 14.5, marginBottom: 8 }}>{g.label}</div>
                {rows.map((r) => (
                  <div key={r.d.id} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: 13, borderTop: `1px solid ${T.paperDim}` }}>
                    <span>{r.d.text}</span>
                    <span style={{ color: T.muted }}>{r.count} {r.count === 1 ? "vez" : "veces"}</span>
                  </div>
                ))}
              </Card>
            );
          })}
        </div>
      )}

      {mode === "pornivel" && (
        <div>
          {niveles.map((nivel) => {
            const rows = roster.dinamicasCatalog
              .map((d) => ({ d, count: doneByLevelAndDinamica[`${nivel}|${d.id}`] || 0 }))
              .filter((r) => r.count > 0)
              .sort((a, b) => b.count - a.count);
            if (rows.length === 0) return null;
            return (
              <Card key={nivel}>
                <div style={{ fontWeight: 800, fontSize: 14.5, marginBottom: 8 }}>{nivel}</div>
                {rows.map((r) => (
                  <div key={r.d.id} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: 13, borderTop: `1px solid ${T.paperDim}` }}>
                    <span>{r.d.text}</span>
                    <span style={{ color: T.muted }}>{r.count} {r.count === 1 ? "vez" : "veces"}</span>
                  </div>
                ))}
              </Card>
            );
          })}
        </div>
      )}

      {mode === "portipo" && (
        <div>
          {roster.tiposClase.length === 0 && <Card style={{ textAlign: "center", color: T.muted }}>Sin tipos de clase todavía — créalos en "Gestionar tipos de clase".</Card>}
          {roster.tiposClase.slice().sort((a, b) => a.numero - b.numero).map((tipo) => (
            <Card key={tipo.numero}>
              <div style={{ fontWeight: 800, fontSize: 14.5, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>Tipo {tipo.numero} <Badge tone="purple">{tipo.nivel || "Iniciación 1"}</Badge></div>
              {tipo.dinamicaIds.length === 0 && <div style={{ fontSize: 12.5, color: T.muted }}>Sin dinámicas asignadas.</div>}
              {tipo.dinamicaIds.map((id) => {
                const cat = roster.dinamicasCatalog.find((c) => c.id === id);
                const count = doneByTipoAndDinamica[`${tipo.numero}|${id}`] || 0;
                return (
                  <div key={id} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: 13, borderTop: `1px solid ${T.paperDim}` }}>
                    <span>{cat ? cat.text : "(eliminada)"}</span>
                    <span style={{ color: T.muted }}>{count} {count === 1 ? "vez" : "veces"}</span>
                  </div>
                );
              })}
            </Card>
          ))}
        </div>
      )}

      {mode === "tipos" && (
        <div>
          {!editingTipo && <button style={btn("primary", { marginBottom: 14 })} onClick={() => startEditTipo(null)}><Plus size={15} /> Nuevo tipo de clase</button>}

          {editingTipo && (
            <Card>
              <div style={{ marginBottom: 10 }}>
                <Label>Número (1–40)</Label>
                <input type="number" style={inputStyle()} value={tipoDraft.numero} onChange={(e) => setTipoDraft({ ...tipoDraft, numero: e.target.value })} disabled={editingTipo !== "new"} />
              </div>
              <div style={{ marginBottom: 10 }}>
                <Label>Nivel — solo se podrá usar en clases de este grupo/nivel</Label>
                <select style={inputStyle()} value={tipoDraft.nivel} onChange={(e) => setTipoDraft({ ...tipoDraft, nivel: e.target.value })}>
                  {ALL_LEVELS.map((lv) => <option key={lv} value={lv}>{lv}</option>)}
                </select>
              </div>
              <Label>Dinámicas</Label>
              {tipoDraft.dinamicaIds.map((id) => {
                const cat = roster.dinamicasCatalog.find((c) => c.id === id);
                return (
                  <div key={id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", fontSize: 13 }}>
                    <span>{cat ? cat.text : "(eliminada)"}</span>
                    <button onClick={() => removeDinFromTipoDraft(id)} style={{ border: "none", background: "transparent", color: T.muted, display: "flex" }}><Trash2 size={13} /></button>
                  </div>
                );
              })}
              <div style={{ display: "flex", gap: 8, margin: "8px 0 16px" }}>
                <input style={inputStyle()} placeholder="Añadir dinámica (nueva o del catálogo)" value={newDinTipo} onChange={(e) => setNewDinTipo(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addDinToTipoDraft()} list="dinamicas-catalog-list-tipo" />
                <datalist id="dinamicas-catalog-list-tipo">{roster.dinamicasCatalog.map((c) => <option key={c.id} value={c.text} />)}</datalist>
                <button style={btn("ghost")} onClick={addDinToTipoDraft}><Plus size={15} /></button>
              </div>
              <Label>Escenas (se cargarán tal cual en cada clase de este tipo)</Label>
              <textarea style={inputStyle({ minHeight: 90, resize: "vertical", marginBottom: 14 })} value={tipoDraft.escenas} onChange={(e) => setTipoDraft({ ...tipoDraft, escenas: e.target.value })} />
              <div style={{ display: "flex", gap: 8 }}>
                <button style={btn("primary", { flex: 1, justifyContent: "center" })} onClick={saveTipoDraft}><Save size={14} /> Guardar</button>
                <button style={btn("ghost")} onClick={() => { setEditingTipo(null); setTipoDraft(null); }}>Cancelar</button>
              </div>
            </Card>
          )}

          {!editingTipo && roster.tiposClase.slice().sort((a, b) => a.numero - b.numero).map((tipo) => (
            <Card key={tipo.numero} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14.5, display: "flex", alignItems: "center", gap: 6 }}>Tipo {tipo.numero} <Badge tone="purple">{tipo.nivel || "Iniciación 1"}</Badge></div>
                <div style={{ fontSize: 12, color: T.muted }}>{tipo.dinamicaIds.length} dinámicas</div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button style={btn("ghost", { padding: "7px 10px" })} onClick={() => startEditTipo(tipo)}><Pencil size={13} /></button>
                <button style={btn("danger", { padding: "7px 10px" })} onClick={() => askConfirm("¿Eliminar este tipo de clase?", () => removeTipoClase(tipo.numero))}><Trash2 size={13} /></button>
              </div>
            </Card>
          ))}
          {!editingTipo && roster.tiposClase.length === 0 && <Card style={{ textAlign: "center", color: T.muted }}>Sin tipos de clase todavía.</Card>}
        </div>
      )}
    </div>
  );
}
