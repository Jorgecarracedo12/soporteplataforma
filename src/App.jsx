import React, { useState, useMemo, useEffect } from "react";
import {
  Search, User, Mail, KeyRound, Paperclip, Clock, Building2,
  Flag, MessageSquare, History, LogOut, Inbox, Eye, EyeOff, Tag, Save,
  TrendingUp, TrendingDown, Minus, Download, FileSpreadsheet, FileText,
  CalendarRange, Users, AlertTriangle, CheckCircle2, Settings2,
  PieChart as PieIcon, BarChart3, Activity, Filter as FilterIcon, X,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";

// ============================================================
// CONEXIÓN A LA API REAL (Fase 6)
// ============================================================
// ⚠️ Reemplaza esta URL por la de tu Worker ya desplegado (la misma que usaste
// para probar los endpoints en las fases anteriores). En la Fase 7, cuando
// unifiquemos frontend + API en un solo proyecto, esto cambiará a una ruta
// relativa ("") porque ambos vivirán en el mismo dominio.
const API_BASE = "https://soporte-estudiantil-api.jorgecarracedo.workers.dev";

async function apiFetch(ruta, opciones = {}) {
  const res = await fetch(API_BASE + ruta, {
    credentials: "include",
    headers: { "Content-Type": "application/json", ...(opciones.headers || {}) },
    ...opciones,
  });
  let data = null;
  try { data = await res.json(); } catch { /* respuesta sin cuerpo */ }
  if (!res.ok) {
    const error = new Error((data && data.error) || "Ocurrió un error inesperado.");
    error.status = res.status;
    error.data = data;
    throw error;
  }
  return data;
}
const apiGet = (ruta) => apiFetch(ruta);
const apiPost = (ruta, body) => apiFetch(ruta, { method: "POST", body: JSON.stringify(body) });
const apiPatch = (ruta, body) => apiFetch(ruta, { method: "PATCH", body: JSON.stringify(body) });
const apiDelete = (ruta) => apiFetch(ruta, { method: "DELETE" });

// Traduce un ticket tal como lo devuelve la API (columnas de la base de datos)
// a la forma interna que ya usan todos los componentes de este archivo.
function ticketDesdeApi(t) {
  return {
    id: t.id,
    nombre: t.nombre,
    codigo: t.identificacion,
    correo: t.correo,
    prizmaUsuario: t.prizma_usuario,
    prizmaPassword: t.prizma_password,
    categoria: t.categoria,
    subcategoria: t.subcategoria,
    descripcion: t.descripcion,
    area: t.area,
    prioridad: t.prioridad,
    estado: t.estado,
    asesor: t.asesor_nombre || "Sin asignar",
    novedadArea: t.novedad_area || "",
    historialRespuestas: (t.respuestas || []).map((r) => ({ fecha: r.fecha, texto: r.texto, autor: "" })),
    comentarios: (t.historial || []).slice().reverse().map((h) => ({ autor: h.autor_nombre || "", texto: h.texto, fecha: h.fecha })),
    creado: t.creado,
    actualizado: t.actualizado,
  };
}
function agenteDesdeApi(a) {
  return { id: a.id, usuario: a.usuario, nombre: a.nombre, rol: a.rol, activo: !!a.activo, avatar: a.avatar, fechaCreacion: a.fecha_creacion };
}

// ---------- Datos de referencia ----------
const CATEGORIAS = ["Plataforma", "Acceso", "Virtual", "Otro"];

// Áreas internas a las que se puede remitir un caso (uso exclusivo del panel de soporte)
const AREAS = ["Prizma", "Partikle", "Problemas LDAP", "Coursera", "Cargue", "Coordinación Académica", "Admisiones"];
const COLOR_AREA = ["#1652F0", "#0BA5EC", "#7C5CFC", "#F59E0B", "#EF4444", "#10B981", "#64748B"];
const COLOR_CATEGORIA = ["#1652F0", "#7C5CFC", "#94A3B8"];

const PRIORIDADES = ["Alta", "Media", "Baja"];
const PRIORIDAD_STYLE = { "Alta": "#D64545", "Media": "#E8A93B", "Baja": "#8FA3BF" };

// Estados del ciclo de vida del ticket, gestionados por el equipo de soporte
const ESTADOS = ["Abierto / Recibido", "En gestión", "Novedad", "Cerrado / Subsanado"];
const ESTADO_ICONO = { "Abierto / Recibido": Clock, "En gestión": Settings2, "Novedad": AlertTriangle, "Cerrado / Subsanado": CheckCircle2 };

// Nota: la lista de asesores ya no es estática. Se deriva dinámicamente del estado
// "agentes" (ver App y DashboardSoporte), para reflejar altas/bajas hechas por el Líder de Área.

// Categorías y subcategorías disponibles para el formulario del estudiante.
// Estructura pensada para poder agregar/quitar categorías o subcategorías fácilmente a futuro.
const CATEGORIAS_ESTUDIANTE = {
  "Plataforma Prizma": [
    "No recuerdo mis credenciales",
    "No sé usar la plataforma",
  ],
  "Inconveniente con Asignaturas": [
    "No visualizo ninguna de mis asignaturas.",
    "Visualizo solo algunas de mis asignaturas.",
    "Las asignaturas que visualizo son incorrectas.",
    "¿Cómo ingreso a Coursera?",
    "¿Cómo ingreso a la plataforma de inglés?",
  ],
  "Otro": [],
};

// Estados que impiden que un estudiante cree un nuevo ticket (ya tiene uno activo).
// Solo "Cerrado / Subsanado" libera al estudiante para crear una nueva solicitud.
const ESTADOS_BLOQUEAN_NUEVO_TICKET = ["Abierto / Recibido", "En gestión", "Novedad"];

const ESTADO_STYLE = {
  "Abierto / Recibido": { bg: "#FDF1DD", fg: "#8A5A0C", dot: "#E8A93B" },
  "En gestión": { bg: "#E9F0FE", fg: "#1D4FB8", dot: "#1652F0" },
  "Novedad": { bg: "#FCEAEA", fg: "#A62E2E", dot: "#D64545" },
  "Cerrado / Subsanado": { bg: "#E6F6ED", fg: "#1D7A47", dot: "#2FAE6B" },
};

// Roles del sistema y sus permisos
const ROL_LIDER = "Líder de Área";
const ROL_ASESOR = "Asesor";
const ROLES = [ROL_LIDER, ROL_ASESOR];

// Avatares predeterminados (sin fotos externas ni marcas registradas)
const AVATARES = ["🧑‍💼", "👩‍💻", "🧑‍🏫", "👨‍🎓", "🦉", "🚀", "🌟", "🎯"];

function nowStr() {
  const d = new Date();
  return d.toISOString().slice(0, 16).replace("T", " ");
}

function diffHoras(fechaInicioStr, fechaFinStr) {
  const a = new Date(fechaInicioStr.replace(" ", "T"));
  const b = new Date(fechaFinStr.replace(" ", "T"));
  return Math.max(0, (b - a) / (1000 * 60 * 60));
}

function formatearFechaCorta(fechaStr) {
  // "2026-07-27 09:12" -> "27/07/2026"
  const soloFecha = fechaStr.slice(0, 10);
  const [anio, mes, dia] = soloFecha.split("-");
  return `${dia}/${mes}/${anio}`;
}

// ---------- Componente: Stub / cupón de ticket ----------
function TicketStub({ id, small }) {
  return (
    <div
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        background: "#1B2430",
        color: "#F5F7FB",
        fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
        fontSize: small ? 12 : 14,
        letterSpacing: "0.04em",
        padding: small ? "3px 8px" : "5px 12px",
        borderRadius: 3,
        position: "relative",
      }}
    >
      <span style={{ opacity: 0.5 }}>#</span>
      <span>TCK-{String(id).padStart(4, "0")}</span>
    </div>
  );
}

function EstadoBadge({ estado }) {
  const s = ESTADO_STYLE[estado];
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        background: s.bg,
        color: s.fg,
        fontSize: 12,
        fontWeight: 600,
        padding: "3px 10px",
        borderRadius: 999,
      }}
    >
      <span style={{ width: 6, height: 6, borderRadius: 999, background: s.dot }} />
      {estado}
    </span>
  );
}

// ---------- Estilos institucionales globales (solo visual, no afecta lógica) ----------
function EstiloGlobalInstitucional() {
  return (
    <style>{`
      button { transition: filter .15s ease, transform .15s ease, box-shadow .15s ease; }
      button:hover { filter: brightness(0.97); }
      button:active { transform: translateY(1px); }
      input, select, textarea { transition: border-color .15s ease, box-shadow .15s ease; }
      input:focus, select:focus, textarea:focus {
        border-color: #1652F0 !important;
        box-shadow: 0 0 0 3px rgba(22, 82, 240, 0.12);
        outline: none;
      }
      [style*="1px solid #E2E6EC"] {
        box-shadow: 0 1px 2px rgba(16, 24, 40, 0.03), 0 1px 3px rgba(16, 24, 40, 0.04);
      }
      a { transition: opacity .15s ease; }
      a:hover { opacity: 0.85; }
    `}</style>
  );
}

// ---------- Vista: Portada ----------
function Portada({ onSelect }) {
  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#F5F7FB" }}>
      {/* Barra superior institucional */}
      <div className="flex items-center justify-center py-3" style={{ background: "#0F1B33" }}>
        <div className="flex items-center gap-2">
          <div style={{ width: 22, height: 22, borderRadius: 6, background: "linear-gradient(135deg, #1652F0, #4C8CFF)" }} />
          <span style={{ color: "#F5F7FB", fontSize: 13, fontWeight: 700, letterSpacing: "0.04em" }}>PRIZMA</span>
          <span style={{ color: "#6E7C99", fontSize: 12 }}>· Ecosistema institucional</span>
        </div>
      </div>

      <div className="flex-1 flex items-center justify-center px-6 py-14">
        <div className="w-full max-w-3xl">
          <div className="text-center mb-10">
            <div
              className="inline-flex items-center gap-2 mb-5 px-3 py-1.5 rounded-full"
              style={{ background: "#EAF0FF", border: "1px solid #D6E2FF" }}
            >
              <span style={{ width: 6, height: 6, borderRadius: 999, background: "#1652F0" }} />
              <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", color: "#1652F0", textTransform: "uppercase" }}>
                Soporte Estudiantil
              </span>
            </div>
            <h1 className="text-4xl sm:text-5xl font-bold" style={{ color: "#0F1B33", letterSpacing: "-0.02em" }}>
              SOPORTE ESTUDIANTIL
            </h1>
            <p className="mt-4 text-base max-w-xl mx-auto" style={{ color: "#5B6472" }}>
              Reporta un inconveniente técnico con Prizma y las plataformas asociadas, o ingresa como
              asesor para gestionar las solicitudes recibidas.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5 mb-8">
            <button
              onClick={() => onSelect("estudiante")}
              className="text-left p-6 rounded-2xl"
              style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}
            >
              <div className="flex items-center justify-between mb-4">
                <div style={{ width: 40, height: 40, borderRadius: 10, background: "#EAF0FF", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <User size={18} color="#1652F0" />
                </div>
                <span style={{ color: "#B7BFCB" }}>→</span>
              </div>
              <span style={{ fontSize: 11, fontWeight: 700, color: "#1652F0", letterSpacing: "0.08em" }}>ESTUDIANTE</span>
              <div className="text-lg font-semibold mt-1 mb-1" style={{ color: "#0F1B33" }}>
                Reportar un problema
              </div>
              <div className="text-sm" style={{ color: "#5B6472" }}>
                Crea un ticket y consulta el estado de tus solicitudes anteriores.
              </div>
            </button>

            <button
              onClick={() => onSelect("soporte")}
              className="text-left p-6 rounded-2xl"
              style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}
            >
              <div className="flex items-center justify-between mb-4">
                <div style={{ width: 40, height: 40, borderRadius: 10, background: "#EEF1F6", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <Inbox size={18} color="#0F1B33" />
                </div>
                <span style={{ color: "#B7BFCB" }}>→</span>
              </div>
              <span style={{ fontSize: 11, fontWeight: 700, color: "#0F1B33", letterSpacing: "0.08em" }}>EQUIPO DE SOPORTE</span>
              <div className="text-lg font-semibold mt-1 mb-1" style={{ color: "#0F1B33" }}>
                Ingresar al panel
              </div>
              <div className="text-sm" style={{ color: "#5B6472" }}>
                Gestiona, redirige y da solución a los tickets recibidos.
              </div>
            </button>
          </div>

          {/* Acceso directo a Prizma */}
          <div className="rounded-2xl p-6 sm:p-7 flex flex-col sm:flex-row items-center justify-between gap-5"
            style={{ background: "linear-gradient(135deg, #0F1B33, #16244A)" }}>
            <div className="text-center sm:text-left">
              <div style={{ color: "#B9C6E6", fontSize: 12, fontWeight: 600, letterSpacing: "0.05em" }}>PLATAFORMA ACADÉMICA</div>
              <div style={{ color: "#FFFFFF", fontSize: 17, fontWeight: 700 }}>¿Deseas ingresar directamente a la plataforma Prizma?</div>
            </div>
            <a href="https://prizma.site/landing" target="_blank" rel="noopener noreferrer"
              className="inline-flex items-center gap-2 px-6 py-3 rounded-xl text-sm font-bold whitespace-nowrap"
              style={{ background: "#1652F0", color: "#FFFFFF" }}>
              Ir a Prizma →
            </a>
          </div>
        </div>
      </div>

      <div className="text-center text-xs pb-6" style={{ color: "#9AA4B2" }}>
        <div className="mb-1">Prototipo de demostración — los datos se reinician al recargar la página.</div>
        <div>Desarrollado por <strong style={{ color: "#5B6472" }}>Jorge Carracedo Cortes</strong></div>
      </div>
    </div>
  );
}

// ---------- Formulario vacío por defecto ----------
const CATEGORIA_INICIAL = Object.keys(CATEGORIAS_ESTUDIANTE)[0];
const FORM_VACIO = {
  nombre: "",
  codigo: "", // internamente sigue siendo "codigo" (en pantalla: "Número de identificación")
  correo: "",
  prizmaUsuario: "",
  prizmaPassword: "",
  categoria: CATEGORIA_INICIAL,
  subcategoria: CATEGORIAS_ESTUDIANTE[CATEGORIA_INICIAL][0] || "",
  descripcion: "",
};

// ---------- Vista: Estudiante ----------
function VistaEstudiante({ onBack }) {
  const [modo, setModo] = useState("crear"); // 'crear' | 'consultar'
  const [form, setForm] = useState(FORM_VACIO);
  const [enviado, setEnviado] = useState(null);
  const [buscarCodigo, setBuscarCodigo] = useState("");
  const [resultados, setResultados] = useState(null);
  const [errorForm, setErrorForm] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [buscando, setBuscando] = useState(false);
  const [ticketActivoExistente, setTicketActivoExistente] = useState(null);

  const subcategoriasDisponibles = CATEGORIAS_ESTUDIANTE[form.categoria] || [];
  const requiereSubcategoria = subcategoriasDisponibles.length > 0;

  // Al salir del campo de identificación, le preguntamos a la API si ya tiene un ticket activo
  async function verificarTicketActivo() {
    const id = form.codigo.trim();
    if (!id) { setTicketActivoExistente(null); return; }
    try {
      const encontrados = await apiGet(`/api/tickets/buscar?identificacion=${encodeURIComponent(id)}`);
      const activo = encontrados.find((t) => ESTADOS_BLOQUEAN_NUEVO_TICKET.includes(t.estado));
      setTicketActivoExistente(activo || null);
    } catch {
      setTicketActivoExistente(null);
    }
  }

  function handleCategoriaChange(categoria) {
    const subs = CATEGORIAS_ESTUDIANTE[categoria] || [];
    setForm({ ...form, categoria, subcategoria: subs[0] || "" });
  }

  async function handleSubmit() {
    // 1) Validar campos obligatorios
    const faltaSubcategoria = requiereSubcategoria && !form.subcategoria;
    if (!form.nombre || !form.codigo || !form.correo || !form.prizmaUsuario || !form.prizmaPassword || !form.descripcion || faltaSubcategoria) {
      setErrorForm("Por favor completa todos los campos obligatorios (*).");
      return;
    }

    setErrorForm("");
    setEnviando(true);
    try {
      const resultado = await apiPost("/api/tickets", {
        nombre: form.nombre.trim(),
        identificacion: form.codigo.trim(),
        correo: form.correo.trim(),
        prizma_usuario: form.prizmaUsuario.trim(),
        prizma_password: form.prizmaPassword,
        categoria: form.categoria,
        subcategoria: form.subcategoria || null,
        descripcion: form.descripcion.trim(),
      });
      setEnviado({ id: resultado.id });
    } catch (err) {
      // La API devuelve 409 cuando ya existe un ticket activo con esa identificación
      setErrorForm(err.message);
      if (err.status === 409) verificarTicketActivo();
    } finally {
      setEnviando(false);
    }
  }

  async function handleBuscar() {
    setBuscando(true);
    setErrorForm("");
    try {
      const encontrados = await apiGet(`/api/tickets/buscar?identificacion=${encodeURIComponent(buscarCodigo.trim())}`);
      setResultados(encontrados.map(ticketDesdeApi));
    } catch (err) {
      setErrorForm(err.message);
      setResultados([]);
    } finally {
      setBuscando(false);
    }
  }


  return (
    <div className="min-h-screen" style={{ background: "#F5F7FB" }}>
      <div className="max-w-2xl mx-auto px-6 py-10">
        <button onClick={onBack} className="text-sm mb-6" style={{ color: "#5A6577" }}>
          ← Volver
        </button>

        <div className="flex gap-2 mb-8">
          <button
            onClick={() => { setModo("crear"); setEnviado(null); }}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{
              background: modo === "crear" ? "#1B2430" : "#FFFFFF",
              color: modo === "crear" ? "#FFFFFF" : "#5A6577",
              border: "1px solid #E2E6EC",
            }}
          >
            Nuevo ticket
          </button>
          <button
            onClick={() => setModo("consultar")}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{
              background: modo === "consultar" ? "#1B2430" : "#FFFFFF",
              color: modo === "consultar" ? "#FFFFFF" : "#5A6577",
              border: "1px solid #E2E6EC",
            }}
          >
            Consultar mis tickets
          </button>
        </div>

        {modo === "crear" && !enviado && (
          <div className="p-6 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
            <h2 className="text-lg font-semibold mb-5" style={{ color: "#1B2430" }}>
              Cuéntanos qué problema presentas
            </h2>

            {ticketActivoExistente && (
              <div className="mb-4 p-3 rounded-lg text-xs" style={{ background: "#FDF1DD", color: "#8A5A0C" }}>
                Ya tienes un ticket activo <strong>TCK-{String(ticketActivoExistente.id).padStart(4, "0")}</strong> con
                estado "{ticketActivoExistente.estado}". Debes esperar a que sea gestionado o cerrado antes de crear uno nuevo.
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>NOMBRE COMPLETO *</label>
                <input value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              </div>
              <div>
                <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>NÚMERO DE IDENTIFICACIÓN *</label>
                <input value={form.codigo} onChange={(e) => setForm({ ...form, codigo: e.target.value })}
                  onBlur={verificarTicketActivo}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              </div>
            </div>

            <div className="mb-4">
              <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>CORREO DEL ESTUDIANTE *</label>
              <input type="email" value={form.correo} onChange={(e) => setForm({ ...form, correo: e.target.value })}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>USUARIO DE PRIZMA *</label>
                <input value={form.prizmaUsuario} onChange={(e) => setForm({ ...form, prizmaUsuario: e.target.value })}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              </div>
              <div>
                <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>CONTRASEÑA DE PRIZMA *</label>
                <input type="password" value={form.prizmaPassword} onChange={(e) => setForm({ ...form, prizmaPassword: e.target.value })}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>CATEGORÍA *</label>
                <select value={form.categoria} onChange={(e) => handleCategoriaChange(e.target.value)}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}>
                  {Object.keys(CATEGORIAS_ESTUDIANTE).map((c) => <option key={c}>{c}</option>)}
                </select>
              </div>
              {requiereSubcategoria && (
                <div>
                  <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>SUBCATEGORÍA *</label>
                  <select value={form.subcategoria} onChange={(e) => setForm({ ...form, subcategoria: e.target.value })}
                    className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}>
                    {subcategoriasDisponibles.map((s) => <option key={s}>{s}</option>)}
                  </select>
                </div>
              )}
            </div>

            <div className="mb-4">
              <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>DESCRIBE TU PROBLEMA *</label>
              <textarea rows={4} value={form.descripcion} onChange={(e) => setForm({ ...form, descripcion: e.target.value })}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}
                placeholder="Ej: No puedo acceder a la plataforma desde ayer, me marca error al iniciar sesión." />
            </div>

            {errorForm && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{errorForm}</div>}

            <button type="button" onClick={handleSubmit} disabled={!!ticketActivoExistente || enviando}
              className="w-full py-3 rounded-lg text-sm font-semibold"
              style={{
                background: (ticketActivoExistente || enviando) ? "#C7CEDA" : "#1652F0",
                color: "#FFFFFF",
                cursor: (ticketActivoExistente || enviando) ? "not-allowed" : "pointer",
              }}>
              {enviando ? "Enviando..." : "Enviar ticket"}
            </button>
          </div>
        )}

        {modo === "crear" && enviado && (
          <div className="p-6 rounded-xl text-center" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
            <div className="text-sm mb-3" style={{ color: "#5A6577" }}>Tu ticket fue registrado</div>
            <div className="flex justify-center mb-4"><TicketStub id={enviado.id} /></div>
            <p className="text-sm mb-1" style={{ color: "#1B2430" }}>Guarda este número para hacer seguimiento.</p>
            <p className="text-sm" style={{ color: "#5A6577" }}>El equipo de soporte revisará tu caso pronto.</p>
            <button onClick={() => { setEnviado(null); setForm(FORM_VACIO); setTicketActivoExistente(null); }}
              className="mt-6 text-sm font-semibold" style={{ color: "#1652F0" }}>
              Registrar otro ticket
            </button>
          </div>
        )}

        {modo === "consultar" && (
          <div>
            <div className="flex gap-2 mb-6">
              <input value={buscarCodigo} onChange={(e) => setBuscarCodigo(e.target.value)}
                placeholder="Ingresa tu número de identificación"
                className="flex-1 px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              <button type="button" onClick={handleBuscar} disabled={buscando} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1B2430", color: "#FFFFFF" }}>
                {buscando ? "Buscando..." : "Buscar"}
              </button>
            </div>

            {resultados && resultados.length === 0 && (
              <div className="text-sm p-8 rounded-xl text-center" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC", color: "#9AA4B2" }}>
                No encontramos tickets con ese número de identificación.
              </div>
            )}

            {resultados && (
              <div className="space-y-3">
                {resultados.map((t) => (
                  <TarjetaTicketEstudiante key={t.id} ticket={t} />
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Tarjeta de ticket para el estudiante (solo información pública) ----------
// Nota de diseño: este componente muestra únicamente los campos que se definan como
// visibles para el estudiante. No incluye notas internas ni datos administrativos.
// Pensado para ampliarse a futuro (ej. añadir "próximo paso" o "tiempo estimado")
// simplemente agregando una fila más dentro de la sección "detalle".
function TarjetaTicketEstudiante({ ticket }) {
  const s = ESTADO_STYLE[ticket.estado] || ESTADO_STYLE["Abierto / Recibido"];
  return (
    <div className="rounded-xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
      <div style={{ height: 4, background: s.dot }} />
      <div className="p-5">
        <div className="flex items-center justify-between mb-3">
          <TicketStub id={ticket.id} small />
          <EstadoBadge estado={ticket.estado} />
        </div>

        <div className="text-xs font-semibold mb-1" style={{ color: "#1652F0", letterSpacing: "0.03em" }}>
          {ticket.categoria}{ticket.subcategoria ? " · " + ticket.subcategoria : ""}
        </div>

        <div className="text-sm mb-3" style={{ color: "#1B2430" }}>{ticket.descripcion}</div>

        {ticket.historialRespuestas && ticket.historialRespuestas.length > 0 && (
          <div className="mt-1 mb-3">
            <div className="text-xs font-semibold mb-2" style={{ color: "#1652F0" }}>Respuestas del equipo de soporte</div>
            <div className="space-y-0">
              {ticket.historialRespuestas.map((r, i) => (
                <div key={i} className="flex gap-2.5">
                  <div className="flex flex-col items-center">
                    <span style={{ width: 8, height: 8, borderRadius: 999, background: "#1652F0", marginTop: 4 }} />
                    {i < ticket.historialRespuestas.length - 1 && (
                      <span style={{ width: 2, flex: 1, background: "#DDE6FA", minHeight: 18 }} />
                    )}
                  </div>
                  <div className="pb-3 flex-1">
                    <div className="text-[11px] font-semibold" style={{ color: "#7A8698" }}>{formatearFechaCorta(r.fecha)}</div>
                    <div className="text-sm p-2.5 rounded-lg mt-0.5" style={{ background: "#EEF3FF", color: "#1D4FB8" }}>{r.texto}</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="text-xs" style={{ color: "#9AA4B2" }}>Creado el {ticket.creado}</div>
      </div>
    </div>
  );
}

// ---------- Vista: Login soporte ----------
function LoginSoporte({ onLogin, onBack }) {
  const [usuario, setUsuario] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [ingresando, setIngresando] = useState(false);

  async function handleSubmit() {
    if (!usuario.trim() || !password.trim()) {
      setError("Ingresa tu usuario y contraseña.");
      return;
    }
    setError("");
    setIngresando(true);
    try {
      const datos = await apiPost("/api/auth/login", { usuario: usuario.trim(), password: password.trim() });
      onLogin(datos);
    } catch (err) {
      setError(err.message);
    } finally {
      setIngresando(false);
    }
  }

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#F5F7FB" }}>
      <div className="flex items-center justify-center py-3" style={{ background: "#0F1B33" }}>
        <div className="flex items-center gap-2">
          <div style={{ width: 22, height: 22, borderRadius: 6, background: "linear-gradient(135deg, #1652F0, #4C8CFF)" }} />
          <span style={{ color: "#F5F7FB", fontSize: 13, fontWeight: 700, letterSpacing: "0.04em" }}>PRIZMA</span>
        </div>
      </div>
      <div className="flex-1 flex items-center justify-center px-6">
        <div className="w-full max-w-sm">
          <button onClick={onBack} className="text-sm mb-6" style={{ color: "#5B6472" }}>← Volver</button>
          <div className="p-7 rounded-2xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: "#EAF0FF", display: "flex", alignItems: "center", justifyContent: "center" }} className="mb-4">
              <Inbox size={18} color="#1652F0" />
            </div>
            <h2 className="text-lg font-semibold mb-1" style={{ color: "#0F1B33" }}>Acceso institucional</h2>
            <p className="text-sm mb-5" style={{ color: "#5B6472" }}>Ingresa con tu usuario asignado por el Líder de Área.</p>
            <div className="mb-3">
              <label className="block text-xs font-semibold mb-1" style={{ color: "#5B6472" }}>USUARIO</label>
              <input value={usuario} onChange={(e) => setUsuario(e.target.value)}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} placeholder="Digite su usuario" />
            </div>
            <div className="mb-2">
              <label className="block text-xs font-semibold mb-1" style={{ color: "#5B6472" }}>CONTRASEÑA</label>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") handleSubmit(); }}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} placeholder="Digite su contraseña" />
            </div>
            {error && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{error}</div>}
            <button type="button" onClick={handleSubmit} disabled={ingresando} className="w-full py-3 rounded-lg text-sm font-semibold mt-3" style={{ background: "#1652F0", color: "#FFFFFF" }}>
              {ingresando ? "Ingresando..." : "Ingresar"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Vista: Dashboard de soporte ----------
function DashboardSoporte({ agenteActual: agenteActualProp, onLogout }) {
  const [tickets, setTickets] = useState([]);
  const [agentes, setAgentes] = useState([]);
  const [seleccionado, setSeleccionado] = useState(null);
  const [cargandoInicial, setCargandoInicial] = useState(true);
  const [errorGeneral, setErrorGeneral] = useState("");

  // Versión siempre actualizada del usuario logueado (por si cambió su avatar en "Mi Perfil")
  const agenteActual = agentes.find((a) => a.id === agenteActualProp.id) || agenteActualProp;
  const esLider = agenteActual.rol === ROL_LIDER;

  const [tab, setTab] = useState("tickets"); // 'tickets' | 'reportes' | 'asesores' | 'perfil'
  const [busqueda, setBusqueda] = useState("");
  const [filtroEstado, setFiltroEstado] = useState("Todos");
  const [filtroArea, setFiltroArea] = useState("Todas");
  const [filtroPrioridad, setFiltroPrioridad] = useState("Todas");
  const [filtroAsesor, setFiltroAsesor] = useState("Todos");
  const [seleccionadoId, setSeleccionadoId] = useState(null);
  const [mostrarPassword, setMostrarPassword] = useState(false);
  const [nuevaRespuesta, setNuevaRespuesta] = useState("");
  const [novedadLocal, setNovedadLocal] = useState("");
  const [guardadoOk, setGuardadoOk] = useState(false);

  async function cargarTickets() {
    try {
      const data = await apiGet("/api/tickets");
      setTickets(data.map(ticketDesdeApi));
      setErrorGeneral("");
    } catch (err) {
      setErrorGeneral(err.message);
    }
  }
  async function cargarAgentes() {
    try {
      const data = await apiGet("/api/agentes");
      setAgentes(data.map(agenteDesdeApi));
    } catch (err) {
      setErrorGeneral(err.message);
    }
  }
  async function cargarDetalle(id) {
    try {
      const data = await apiGet(`/api/tickets/${id}`);
      setSeleccionado(ticketDesdeApi(data));
    } catch (err) {
      setErrorGeneral(err.message);
    }
  }

  // Carga inicial
  useEffect(() => {
    (async () => {
      await Promise.all([cargarTickets(), cargarAgentes()]);
      setCargandoInicial(false);
    })();
  }, []);

  // "Tiempo real" simplificado: refrescamos la bandeja cada 8 segundos
  useEffect(() => {
    const intervalo = setInterval(cargarTickets, 8000);
    return () => clearInterval(intervalo);
  }, []);

  // Asesores activos disponibles para asignar/filtrar (se actualiza si el Líder crea, edita o desactiva cuentas)
  const asesoresActivos = useMemo(
    () => agentes.filter((a) => a.rol === ROL_ASESOR && a.activo).map((a) => a.nombre),
    [agentes]
  );

  const filtrados = useMemo(() => {
    const q = busqueda.trim().toLowerCase();
    return tickets.filter((t) => {
      if (filtroEstado !== "Todos" && t.estado !== filtroEstado) return false;
      if (filtroArea !== "Todas" && t.area !== filtroArea) return false;
      if (filtroPrioridad !== "Todas" && t.prioridad !== filtroPrioridad) return false;
      if (filtroAsesor !== "Todos" && t.asesor !== filtroAsesor) return false;
      if (q && !(t.nombre.toLowerCase().includes(q) || t.codigo.toLowerCase().includes(q))) return false;
      return true;
    });
  }, [tickets, filtroEstado, filtroArea, filtroPrioridad, filtroAsesor, busqueda]);

  // Al cambiar de ticket seleccionado, traer su detalle completo y sincronizar los editores locales
  useEffect(() => {
    setNuevaRespuesta("");
    setNovedadLocal("");
    setMostrarPassword(false);
    setGuardadoOk(false);
    setConfirmarEliminar(false);
    if (seleccionadoId) cargarDetalle(seleccionadoId);
    else setSeleccionado(null);
  }, [seleccionadoId]);

  useEffect(() => {
    setNovedadLocal(seleccionado ? seleccionado.novedadArea || "" : "");
  }, [seleccionado]);

  async function actualizarCampo(id, campo, valor, etiqueta) {
    const payload = {};
    if (campo === "asesor") {
      const encontrado = agentes.find((a) => a.nombre === valor);
      payload.asesor_id = valor === "Sin asignar" || !encontrado ? null : encontrado.id;
    } else {
      payload[campo] = valor;
    }
    try {
      await apiPatch(`/api/tickets/${id}`, payload);
      await cargarTickets();
      if (seleccionadoId === id) await cargarDetalle(id);
    } catch (err) {
      setErrorGeneral(err.message);
    }
  }

  async function agregarRespuesta() {
    if (!seleccionado || !nuevaRespuesta.trim()) return;
    try {
      await apiPost(`/api/tickets/${seleccionado.id}/respuestas`, { texto: nuevaRespuesta.trim() });
      setNuevaRespuesta("");
      await cargarDetalle(seleccionado.id);
      await cargarTickets();
    } catch (err) {
      setErrorGeneral(err.message);
    }
  }

  async function guardarGestion() {
    if (!seleccionado) return;
    try {
      await apiPatch(`/api/tickets/${seleccionado.id}`, { novedad_area: novedadLocal });
      await cargarDetalle(seleccionado.id);
      setGuardadoOk(true);
      setTimeout(() => setGuardadoOk(false), 2000);
    } catch (err) {
      setErrorGeneral(err.message);
    }
  }

  const [confirmarEliminar, setConfirmarEliminar] = useState(false);
  const [eliminando, setEliminando] = useState(false);

  async function eliminarTicket() {
    if (!seleccionado) return;
    setEliminando(true);
    try {
      await apiDelete(`/api/tickets/${seleccionado.id}`);
      setSeleccionadoId(null);
      setConfirmarEliminar(false);
      await cargarTickets();
    } catch (err) {
      setErrorGeneral(err.message);
    } finally {
      setEliminando(false);
    }
  }

  return (
    <div className="min-h-screen" style={{ background: "#F1F3F7" }}>
      <style>{`
        .dash-grid { display: grid; grid-template-columns: 320px 1fr 380px; gap: 16px; align-items: start; }
        @media (max-width: 1100px) {
          .dash-grid { grid-template-columns: 1fr; }
        }
        .dash-scroll::-webkit-scrollbar { width: 6px; }
        .dash-scroll::-webkit-scrollbar-thumb { background: #D7DCE3; border-radius: 999px; }
      `}</style>

      {/* Header corporativo */}
      <div className="flex items-center justify-between px-6 py-3" style={{ background: "#0F1B33", borderBottom: "1px solid #232C40" }}>
        <div className="flex items-center gap-3">
          <div style={{ width: 30, height: 30, borderRadius: 8, background: "#1652F0", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Inbox size={16} color="#FFFFFF" />
          </div>
          <div>
            <div style={{ color: "#F5F7FB", fontWeight: 700, fontSize: 14, lineHeight: 1.1 }}>SOPORTE ESTUDIANTIL</div>
            <div style={{ color: "#7A8698", fontSize: 11 }}>Prizma · Panel de gestión de tickets</div>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-lg" style={{ background: "#1C2536" }}>
            <div style={{ width: 22, height: 22, borderRadius: 999, background: "#1652F0", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>
              {agenteActual.avatar || agenteActual.nombre.split(" ").map((p) => p[0]).slice(0, 2).join("")}
            </div>
            <div className="leading-tight">
              <div style={{ color: "#E5E9F0", fontSize: 13 }}>{agenteActual.nombre}</div>
              <div style={{ color: "#7A8698", fontSize: 10 }}>{agenteActual.rol}</div>
            </div>
          </div>
          <button onClick={onLogout} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold"
            style={{ color: "#F5F7FB", background: "#2A3542" }}>
            <LogOut size={13} /> Salir
          </button>
        </div>
      </div>

      <div className="px-4 sm:px-6 py-5" style={{ maxWidth: 1600, margin: "0 auto" }}>
        {errorGeneral && (
          <div className="mb-4 p-3 rounded-lg text-xs" style={{ background: "#FCEAEA", color: "#A62E2E" }}>
            {errorGeneral}
          </div>
        )}
        {cargandoInicial && (
          <div className="mb-4 p-3 rounded-lg text-xs" style={{ background: "#F4F6F9", color: "#5A6577" }}>
            Cargando información...
          </div>
        )}
        <div className="flex gap-2 mb-5 flex-wrap">
          <button onClick={() => setTab("tickets")}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{ background: tab === "tickets" ? "#0F1B33" : "#FFFFFF", color: tab === "tickets" ? "#FFFFFF" : "#5A6577", border: "1px solid #E2E6EC" }}>
            Bandeja de tickets
          </button>
          <button onClick={() => setTab("reportes")}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{ background: tab === "reportes" ? "#0F1B33" : "#FFFFFF", color: tab === "reportes" ? "#FFFFFF" : "#5A6577", border: "1px solid #E2E6EC" }}>
            Reportes
          </button>
          {esLider && (
            <button onClick={() => setTab("asesores")}
              className="px-4 py-2 rounded-lg text-sm font-semibold"
              style={{ background: tab === "asesores" ? "#0F1B33" : "#FFFFFF", color: tab === "asesores" ? "#FFFFFF" : "#5A6577", border: "1px solid #E2E6EC" }}>
              Asesores
            </button>
          )}
          <button onClick={() => setTab("perfil")}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{ background: tab === "perfil" ? "#0F1B33" : "#FFFFFF", color: tab === "perfil" ? "#FFFFFF" : "#5A6577", border: "1px solid #E2E6EC" }}>
            Mi Perfil
          </button>
        </div>

        {tab === "tickets" && (
          <div className="dash-grid">
            {/* ---------- Panel izquierdo: lista ---------- */}
            <div className="rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC", overflow: "hidden" }}>
              <div className="p-3" style={{ borderBottom: "1px solid #EEF0F3" }}>
                <div className="flex items-center gap-2 px-3 py-2 rounded-lg mb-2" style={{ background: "#F4F6F9" }}>
                  <Search size={14} color="#9AA4B2" />
                  <input value={busqueda} onChange={(e) => setBusqueda(e.target.value)}
                    placeholder="Buscar por nombre o identificación"
                    className="flex-1 bg-transparent text-xs outline-none" />
                </div>
                <div className="grid grid-cols-2 gap-1.5">
                  <select value={filtroEstado} onChange={(e) => setFiltroEstado(e.target.value)}
                    className="px-1.5 py-1.5 rounded-lg text-xs" style={{ border: "1px solid #E2E6EC" }}>
                    <option>Todos</option>
                    {ESTADOS.map((e) => <option key={e}>{e}</option>)}
                  </select>
                  <select value={filtroArea} onChange={(e) => setFiltroArea(e.target.value)}
                    className="px-1.5 py-1.5 rounded-lg text-xs" style={{ border: "1px solid #E2E6EC" }}>
                    <option>Todas</option>
                    {AREAS.map((a) => <option key={a}>{a}</option>)}
                  </select>
                  <select value={filtroPrioridad} onChange={(e) => setFiltroPrioridad(e.target.value)}
                    className="px-1.5 py-1.5 rounded-lg text-xs" style={{ border: "1px solid #E2E6EC" }}>
                    <option>Todas</option>
                    {PRIORIDADES.map((p) => <option key={p}>{p}</option>)}
                  </select>
                  <select value={filtroAsesor} onChange={(e) => setFiltroAsesor(e.target.value)}
                    className="px-1.5 py-1.5 rounded-lg text-xs" style={{ border: "1px solid #E2E6EC" }}>
                    <option>Todos</option>
                    <option>Sin asignar</option>
                    {asesoresActivos.map((a) => <option key={a}>{a}</option>)}
                  </select>
                </div>
              </div>

              <div className="dash-scroll" style={{ maxHeight: "calc(100vh - 260px)", overflowY: "auto" }}>
                {filtrados.length === 0 && (
                  <div className="text-xs p-6 text-center" style={{ color: "#9AA4B2" }}>No hay tickets con estos filtros.</div>
                )}
                {filtrados.map((t) => {
                  const s = ESTADO_STYLE[t.estado] || ESTADO_STYLE["Abierto / Recibido"];
                  const activo = seleccionadoId === t.id;
                  return (
                    <button key={t.id} onClick={() => setSeleccionadoId(t.id)}
                      className="w-full text-left px-4 py-3"
                      style={{
                        background: activo ? "#EEF3FF" : "#FFFFFF",
                        borderLeft: `3px solid ${s.dot}`,
                        borderBottom: "1px solid #F0F2F5",
                      }}>
                      <div className="flex items-center justify-between mb-1">
                        <span className="text-sm font-semibold truncate" style={{ color: "#1B2430", maxWidth: 160 }}>{t.nombre}</span>
                        <span style={{ width: 7, height: 7, borderRadius: 999, background: PRIORIDAD_STYLE[t.prioridad] }} title={"Prioridad " + t.prioridad} />
                      </div>
                      <div className="text-xs mb-1.5" style={{ color: "#7A8698" }}>{t.codigo}</div>
                      <div className="text-xs mb-2 truncate" style={{ color: "#5A6577" }}>{t.categoria}</div>
                      <div className="flex items-center justify-between">
                        <EstadoBadge estado={t.estado} />
                        <span className="text-[10px]" style={{ color: "#B7BFCB" }}>{t.creado.slice(0, 10)}</span>
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* ---------- Panel central: información del estudiante ---------- */}
            <div className="rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC", overflow: "hidden" }}>
              {!seleccionado && (
                <div className="p-14 text-center text-sm" style={{ color: "#9AA4B2" }}>
                  Selecciona un ticket de la lista para ver el detalle.
                </div>
              )}
              {seleccionado && (
                <div className="dash-scroll" style={{ maxHeight: "calc(100vh - 220px)", overflowY: "auto" }}>
                  <div className="p-5" style={{ borderBottom: "1px solid #EEF0F3" }}>
                    <div className="flex items-center justify-between mb-1">
                      <TicketStub id={seleccionado.id} />
                      <EstadoBadge estado={seleccionado.estado} />
                    </div>
                    <div className="text-xs mt-2" style={{ color: "#9AA4B2" }}>Creado el {seleccionado.creado}</div>
                  </div>

                  {/* Datos del estudiante */}
                  <div className="p-5" style={{ borderBottom: "1px solid #EEF0F3" }}>
                    <div className="text-xs font-bold mb-3" style={{ color: "#7A8698", letterSpacing: "0.05em" }}>DATOS DEL ESTUDIANTE</div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div className="flex items-start gap-2">
                        <User size={14} color="#9AA4B2" className="mt-0.5" />
                        <div>
                          <div className="text-[11px]" style={{ color: "#9AA4B2" }}>Nombre completo</div>
                          <div className="text-sm font-medium" style={{ color: "#1B2430" }}>{seleccionado.nombre}</div>
                        </div>
                      </div>
                      <div className="flex items-start gap-2">
                        <Tag size={14} color="#9AA4B2" className="mt-0.5" />
                        <div>
                          <div className="text-[11px]" style={{ color: "#9AA4B2" }}>Número de identificación</div>
                          <div className="text-sm font-medium" style={{ color: "#1B2430" }}>{seleccionado.codigo}</div>
                        </div>
                      </div>
                      <div className="flex items-start gap-2 sm:col-span-2">
                        <Mail size={14} color="#9AA4B2" className="mt-0.5" />
                        <div>
                          <div className="text-[11px]" style={{ color: "#9AA4B2" }}>Correo del estudiante</div>
                          <div className="text-sm font-medium" style={{ color: "#1B2430" }}>{seleccionado.correo}</div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Credenciales Prizma */}
                  <div className="p-5" style={{ borderBottom: "1px solid #EEF0F3" }}>
                    <div className="text-xs font-bold mb-3" style={{ color: "#7A8698", letterSpacing: "0.05em" }}>ACCESO A PRIZMA</div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div className="flex items-start gap-2">
                        <KeyRound size={14} color="#9AA4B2" className="mt-0.5" />
                        <div>
                          <div className="text-[11px]" style={{ color: "#9AA4B2" }}>Usuario Prizma</div>
                          <div className="text-sm font-medium" style={{ color: "#1B2430" }}>{seleccionado.prizmaUsuario}</div>
                        </div>
                      </div>
                      <div className="flex items-start gap-2">
                        <KeyRound size={14} color="#9AA4B2" className="mt-0.5" />
                        <div className="flex-1">
                          <div className="text-[11px]" style={{ color: "#9AA4B2" }}>Contraseña Prizma</div>
                          <div className="flex items-center gap-2">
                            <span className="text-sm font-medium" style={{ color: "#1B2430" }}>
                              {mostrarPassword ? seleccionado.prizmaPassword : "•".repeat(Math.max(seleccionado.prizmaPassword.length, 6))}
                            </span>
                            <button onClick={() => setMostrarPassword((v) => !v)} style={{ color: "#1652F0" }}>
                              {mostrarPassword ? <EyeOff size={13} /> : <Eye size={13} />}
                            </button>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Categoría y descripción */}
                  <div className="p-5" style={{ borderBottom: "1px solid #EEF0F3" }}>
                    <div className="text-xs font-bold mb-3" style={{ color: "#7A8698", letterSpacing: "0.05em" }}>DETALLE DEL PROBLEMA</div>
                    <div className="flex items-center gap-2 mb-3">
                      <span className="text-xs font-semibold px-2.5 py-1 rounded-full" style={{ background: "#EEF3FF", color: "#1D4FB8" }}>
                        {seleccionado.categoria}
                      </span>
                      {seleccionado.subcategoria && (
                        <span className="text-xs px-2.5 py-1 rounded-full" style={{ background: "#F4F6F9", color: "#5A6577" }}>
                          {seleccionado.subcategoria}
                        </span>
                      )}
                    </div>
                    <div className="text-sm p-3 rounded-lg" style={{ background: "#F5F7FB", color: "#1B2430" }}>
                      {seleccionado.descripcion}
                    </div>
                  </div>

                  {/* Historial */}
                  <div className="p-5">
                    <div className="flex items-center gap-1.5 text-xs font-bold mb-3" style={{ color: "#7A8698", letterSpacing: "0.05em" }}>
                      <History size={13} /> HISTORIAL DEL TICKET
                    </div>
                    {seleccionado.comentarios.length === 0 && (
                      <div className="text-xs" style={{ color: "#B7BFCB" }}>Aún no hay movimientos registrados.</div>
                    )}
                    <div className="space-y-2">
                      {seleccionado.comentarios.slice().reverse().map((c, i) => (
                        <div key={i} className="text-xs p-2.5 rounded-lg" style={{ background: "#F5F7FB" }}>
                          <div className="flex justify-between mb-0.5">
                            <span className="font-semibold" style={{ color: "#1B2430" }}>{c.autor}</span>
                            <span style={{ color: "#B7BFCB" }}>{c.fecha}</span>
                          </div>
                          <div style={{ color: "#5A6577" }}>{c.texto}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* ---------- Panel derecho: gestión interna ---------- */}
            <div className="rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC", overflow: "hidden" }}>
              {!seleccionado && (
                <div className="p-10 text-center text-sm" style={{ color: "#9AA4B2" }}>
                  Aquí aparecerán los controles de gestión.
                </div>
              )}
              {seleccionado && (
                <div className="dash-scroll" style={{ maxHeight: "calc(100vh - 220px)", overflowY: "auto" }}>
                  <div className="px-5 py-4" style={{ background: "#0F1B33" }}>
                    <div className="text-xs font-bold" style={{ color: "#F5F7FB", letterSpacing: "0.05em" }}>GESTIÓN INTERNA</div>
                    <div className="text-[11px]" style={{ color: "#8391A6" }}>Esta información no es visible para el estudiante</div>
                  </div>

                  <div className="p-5 space-y-4" style={{ borderBottom: "1px solid #EEF0F3" }}>
                    <div>
                      <label className="flex items-center gap-1.5 text-[11px] font-semibold mb-1" style={{ color: "#5A6577" }}>
                        <Flag size={12} /> ESTADO
                      </label>
                      <select value={seleccionado.estado} onChange={(e) => actualizarCampo(seleccionado.id, "estado", e.target.value, "Estado")}
                        className="w-full px-2.5 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}>
                        {ESTADOS.map((e) => <option key={e}>{e}</option>)}
                      </select>
                    </div>

                    <div>
                      <label className="flex items-center gap-1.5 text-[11px] font-semibold mb-1" style={{ color: "#5A6577" }}>
                        <Building2 size={12} /> ÁREA
                      </label>
                      <select value={seleccionado.area} onChange={(e) => actualizarCampo(seleccionado.id, "area", e.target.value, "Área")}
                        className="w-full px-2.5 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}>
                        <option>Sin asignar</option>
                        {AREAS.map((a) => <option key={a}>{a}</option>)}
                      </select>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>PRIORIDAD</label>
                        <select value={seleccionado.prioridad} onChange={(e) => actualizarCampo(seleccionado.id, "prioridad", e.target.value, "Prioridad")}
                          className="w-full px-2.5 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}>
                          {PRIORIDADES.map((p) => <option key={p}>{p}</option>)}
                        </select>
                      </div>
                      <div>
                        <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>ASESOR</label>
                        <select value={seleccionado.asesor} onChange={(e) => actualizarCampo(seleccionado.id, "asesor", e.target.value, "Asesor")}
                          className="w-full px-2.5 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }}>
                          <option>Sin asignar</option>
                          {asesoresActivos.map((a) => <option key={a}>{a}</option>)}
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="p-5 space-y-4" style={{ borderBottom: "1px solid #EEF0F3" }}>
                    <div>
                      <label className="flex items-center gap-1.5 text-[11px] font-semibold mb-1" style={{ color: "#5A6577" }}>
                        <MessageSquare size={12} /> RESPUESTA AL ESTUDIANTE
                      </label>
                      <div className="text-[11px] mb-2" style={{ color: "#9AA4B2" }}>
                        Cada mensaje se agrega al historial; el estudiante ve todas las respuestas en orden cronológico.
                      </div>

                      {seleccionado.historialRespuestas && seleccionado.historialRespuestas.length > 0 && (
                        <div className="space-y-1.5 mb-2" style={{ maxHeight: 140, overflowY: "auto" }}>
                          {seleccionado.historialRespuestas.slice().reverse().map((r, i) => (
                            <div key={i} className="text-xs p-2 rounded-lg" style={{ background: "#EEF3FF" }}>
                              <div className="font-semibold" style={{ color: "#1D4FB8" }}>{formatearFechaCorta(r.fecha)}</div>
                              <div style={{ color: "#1B2430" }}>{r.texto}</div>
                            </div>
                          ))}
                        </div>
                      )}

                      <div className="flex gap-2">
                        <textarea rows={2} value={nuevaRespuesta} onChange={(e) => setNuevaRespuesta(e.target.value)}
                          placeholder="Escribe una nueva respuesta para el estudiante..."
                          className="flex-1 px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                      </div>
                      <button onClick={agregarRespuesta} disabled={!nuevaRespuesta.trim()}
                        className="mt-2 w-full flex items-center justify-center gap-2 py-2 rounded-lg text-sm font-semibold"
                        style={{ background: nuevaRespuesta.trim() ? "#1B2430" : "#E2E6EC", color: nuevaRespuesta.trim() ? "#FFFFFF" : "#9AA4B2" }}>
                        <MessageSquare size={13} /> Agregar respuesta
                      </button>
                    </div>
                  </div>

                  <div className="p-5 space-y-4">
                    <div>
                      <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>NOVEDAD ÁREA</label>
                      <div className="text-[11px] mb-1.5" style={{ color: "#9AA4B2" }}>Nota interna para remitir el caso entre áreas. No visible para el estudiante.</div>
                      <textarea rows={3} value={novedadLocal} onChange={(e) => setNovedadLocal(e.target.value)}
                        placeholder="Ej: Se remite a Prizma para validar el bloqueo de la cuenta."
                        className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                    </div>

                    <button onClick={guardarGestion}
                      className="w-full flex items-center justify-center gap-2 py-2.5 rounded-lg text-sm font-semibold"
                      style={{ background: guardadoOk ? "#2FAE6B" : "#1652F0", color: "#FFFFFF" }}>
                      <Save size={14} /> {guardadoOk ? "Guardado" : "Guardar novedad de área"}
                    </button>
                  </div>

                  <div className="p-5" style={{ borderTop: "1px solid #EEF0F3" }}>
                    <div className="text-[11px] font-bold mb-2" style={{ color: "#A62E2E", letterSpacing: "0.05em" }}>ZONA DE PELIGRO</div>
                    {!confirmarEliminar ? (
                      <button onClick={() => setConfirmarEliminar(true)}
                        className="w-full flex items-center justify-center gap-2 py-2.5 rounded-lg text-sm font-semibold"
                        style={{ background: "#FCEAEA", color: "#A62E2E" }}>
                        Eliminar ticket
                      </button>
                    ) : (
                      <div className="p-3 rounded-lg" style={{ background: "#FCEAEA" }}>
                        <div className="text-xs mb-3" style={{ color: "#A62E2E" }}>
                          Esta acción es permanente y borrará el ticket, sus respuestas e historial. ¿Confirmas?
                        </div>
                        <div className="flex gap-2">
                          <button onClick={eliminarTicket} disabled={eliminando}
                            className="flex-1 py-2 rounded-lg text-xs font-semibold" style={{ background: "#D64545", color: "#FFFFFF" }}>
                            {eliminando ? "Eliminando..." : "Sí, eliminar"}
                          </button>
                          <button onClick={() => setConfirmarEliminar(false)}
                            className="flex-1 py-2 rounded-lg text-xs font-semibold" style={{ background: "#FFFFFF", color: "#5A6577" }}>
                            Cancelar
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "reportes" && <ModuloReportes tickets={tickets} agentes={agentes} agenteActual={agenteActual} />}

        {tab === "asesores" && esLider && (
          <PanelAsesores agentes={agentes} onCambio={cargarAgentes} agenteActual={agenteActual} />
        )}

        {tab === "perfil" && (
          <PanelMiPerfil agenteActual={agenteActual} onCambio={cargarAgentes} />
        )}
      </div>
    </div>
  );
}

// ---------- Tarjeta KPI ----------
function TarjetaKPI({ titulo, valor, porcentaje, variacion, color, Icono }) {
  return (
    <div className="p-4 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
      <div className="flex items-center justify-between mb-3">
        <div style={{ width: 32, height: 32, borderRadius: 8, background: color + "1A", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icono size={16} color={color} />
        </div>
        {variacion !== null && variacion !== undefined ? (
          <span className="flex items-center gap-1 text-xs font-semibold" style={{ color: variacion > 0 ? "#2FAE6B" : variacion < 0 ? "#D64545" : "#9AA4B2" }}>
            {variacion > 0 ? <TrendingUp size={12} /> : variacion < 0 ? <TrendingDown size={12} /> : <Minus size={12} />}
            {Math.abs(variacion)}%
          </span>
        ) : (
          <span className="text-[10px]" style={{ color: "#C7CEDA" }}>N/D</span>
        )}
      </div>
      <div className="text-2xl font-bold" style={{ color: "#1B2430" }}>{valor}</div>
      <div className="text-xs mt-0.5" style={{ color: "#9AA4B2" }}>{titulo} · {porcentaje}%</div>
    </div>
  );
}

// ---------- Módulo: Reportes para el perfil Asesor (vista restringida) ----------
// El asesor solo ve: un resumen general (totales por estado, sin desagregar por otros
// asesores) y su propia gestión personal. No tiene filtros globales ni exportación.
function ModuloReportesAsesor({ tickets, agenteActual }) {
  const resumenGeneral = useMemo(() => {
    const porEstado = {};
    ESTADOS.forEach((e) => (porEstado[e] = 0));
    tickets.forEach((t) => { porEstado[t.estado] = (porEstado[t.estado] || 0) + 1; });
    return { total: tickets.length, porEstado };
  }, [tickets]);

  const miGestion = useMemo(() => {
    const propios = tickets.filter((t) => t.asesor === agenteActual.nombre);
    const cerrados = propios.filter((t) => t.estado === "Cerrado / Subsanado");
    let tiempoPromedio = null;
    if (cerrados.length > 0) {
      const totalHoras = cerrados.reduce((s, t) => s + diffHoras(t.creado, t.actualizado), 0);
      tiempoPromedio = Math.round((totalHoras / cerrados.length) * 10) / 10;
    }
    return {
      asignados: propios.length,
      gestionados: propios.filter((t) => t.estado !== "Abierto / Recibido").length,
      cerrados: cerrados.length,
      pendientes: propios.filter((t) => t.estado === "Abierto / Recibido").length,
      tiempoPromedio,
    };
  }, [tickets, agenteActual]);

  return (
    <div>
      <div className="p-5 rounded-xl mb-6" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="text-sm font-semibold mb-1" style={{ color: "#1B2430" }}>Resumen general de la operación</div>
        <div className="text-xs mb-4" style={{ color: "#9AA4B2" }}>
          Información informativa de toda la mesa de ayuda. No incluye el detalle de gestión de otros asesores.
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-4">
          <div className="p-3 rounded-lg" style={{ background: "#F5F7FB" }}>
            <div className="text-xl font-bold" style={{ color: "#1B2430" }}>{resumenGeneral.total}</div>
            <div className="text-xs" style={{ color: "#9AA4B2" }}>Total tickets</div>
          </div>
          {ESTADOS.map((e) => (
            <div key={e} className="p-3 rounded-lg" style={{ background: "#F5F7FB" }}>
              <div className="text-xl font-bold" style={{ color: ESTADO_STYLE[e].dot }}>{resumenGeneral.porEstado[e]}</div>
              <div className="text-xs" style={{ color: "#9AA4B2" }}>{e}</div>
            </div>
          ))}
        </div>
      </div>

      <div className="p-5 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
          <Users size={15} color="#1652F0" /> Mi gestión
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          <div className="p-3 rounded-lg" style={{ background: "#EEF3FF" }}>
            <div className="text-xl font-bold" style={{ color: "#1D4FB8" }}>{miGestion.asignados}</div>
            <div className="text-xs" style={{ color: "#5A6577" }}>Asignados</div>
          </div>
          <div className="p-3 rounded-lg" style={{ background: "#F4F6F9" }}>
            <div className="text-xl font-bold" style={{ color: "#1B2430" }}>{miGestion.gestionados}</div>
            <div className="text-xs" style={{ color: "#5A6577" }}>Gestionados</div>
          </div>
          <div className="p-3 rounded-lg" style={{ background: "#E6F6ED" }}>
            <div className="text-xl font-bold" style={{ color: "#1D7A47" }}>{miGestion.cerrados}</div>
            <div className="text-xs" style={{ color: "#5A6577" }}>Cerrados</div>
          </div>
          <div className="p-3 rounded-lg" style={{ background: "#FDF1DD" }}>
            <div className="text-xl font-bold" style={{ color: "#8A5A0C" }}>{miGestion.pendientes}</div>
            <div className="text-xs" style={{ color: "#5A6577" }}>Pendientes</div>
          </div>
        </div>
        <div className="text-xs mt-3" style={{ color: "#9AA4B2" }}>
          Tiempo promedio de gestión: {miGestion.tiempoPromedio !== null ? `${miGestion.tiempoPromedio} h` : "Sin datos suficientes"}
        </div>
      </div>
    </div>
  );
}

// ---------- Módulo: Dashboard Ejecutivo de Reportes ----------
function ModuloReportes({ tickets, agentes, agenteActual }) {
  const categorias = Object.keys(CATEGORIAS_ESTUDIANTE);
  // Todos los asesores que existen (activos o no), para que el histórico de reportes no pierda datos
  const asesoresReales = useMemo(
    () => agentes.filter((a) => a.rol === ROL_ASESOR).map((a) => a.nombre),
    [agentes]
  );

  if (agenteActual.rol === ROL_ASESOR) {
    return <ModuloReportesAsesor tickets={tickets} agenteActual={agenteActual} />;
  }

  // ----- Filtros -----
  const [fInicio, setFInicio] = useState("");
  const [fFin, setFFin] = useState("");
  const [fEstado, setFEstado] = useState("Todos");
  const [fArea, setFArea] = useState("Todas");
  const [fAsesor, setFAsesor] = useState("Todos");
  const [fCategoria, setFCategoria] = useState("Todas");
  const [fPrioridad, setFPrioridad] = useState("Todas");

  function limpiarFiltros() {
    setFInicio(""); setFFin(""); setFEstado("Todos"); setFArea("Todas");
    setFAsesor("Todos"); setFCategoria("Todas"); setFPrioridad("Todas");
  }
  const hayFiltrosActivos = fInicio || fFin || fEstado !== "Todos" || fArea !== "Todas" || fAsesor !== "Todos" || fCategoria !== "Todas" || fPrioridad !== "Todas";

  function cumpleFiltrosBase(t) {
    if (fEstado !== "Todos" && t.estado !== fEstado) return false;
    if (fArea !== "Todas" && t.area !== fArea) return false;
    if (fAsesor !== "Todos" && t.asesor !== fAsesor) return false;
    if (fCategoria !== "Todas" && t.categoria !== fCategoria) return false;
    if (fPrioridad !== "Todas" && t.prioridad !== fPrioridad) return false;
    return true;
  }

  const ticketsFiltrados = useMemo(() => {
    return tickets.filter((t) => {
      const fecha = t.creado.slice(0, 10);
      if (fInicio && fecha < fInicio) return false;
      if (fFin && fecha > fFin) return false;
      return cumpleFiltrosBase(t);
    });
  }, [tickets, fInicio, fFin, fEstado, fArea, fAsesor, fCategoria, fPrioridad]);

  // ----- KPIs -----
  const kpis = useMemo(() => {
    const porEstado = {};
    ESTADOS.forEach((e) => (porEstado[e] = 0));
    ticketsFiltrados.forEach((t) => { porEstado[t.estado] = (porEstado[t.estado] || 0) + 1; });
    return { total: ticketsFiltrados.length, porEstado };
  }, [ticketsFiltrados]);

  // Variación respecto al período anterior de igual longitud (solo si hay rango de fechas)
  const variacion = useMemo(() => {
    if (!fInicio || !fFin) return null;
    const d1 = new Date(fInicio + "T00:00:00");
    const d2 = new Date(fFin + "T00:00:00");
    const dias = Math.max(1, Math.round((d2 - d1) / 86400000) + 1);
    const finAnterior = new Date(d1); finAnterior.setDate(finAnterior.getDate() - 1);
    const inicioAnterior = new Date(finAnterior); inicioAnterior.setDate(inicioAnterior.getDate() - dias + 1);
    const fmt = (d) => d.toISOString().slice(0, 10);
    const anteriores = tickets.filter((t) => {
      const fecha = t.creado.slice(0, 10);
      return fecha >= fmt(inicioAnterior) && fecha <= fmt(finAnterior) && cumpleFiltrosBase(t);
    });
    const pct = (curr, prev) => (prev === 0 ? (curr === 0 ? 0 : 100) : Math.round(((curr - prev) / prev) * 100));
    const out = { total: pct(kpis.total, anteriores.length) };
    ESTADOS.forEach((e) => { out[e] = pct(kpis.porEstado[e], anteriores.filter((t) => t.estado === e).length); });
    return out;
  }, [tickets, fInicio, fFin, fEstado, fArea, fAsesor, fCategoria, fPrioridad, kpis]);

  // ----- Gestión por asesor -----
  const gestionAsesores = useMemo(() => {
    return asesoresReales.map((a) => {
      const propios = ticketsFiltrados.filter((t) => t.asesor === a);
      const cerrados = propios.filter((t) => t.estado === "Cerrado / Subsanado");
      let tiempoPromedio = null;
      if (cerrados.length > 0) {
        const totalHoras = cerrados.reduce((s, t) => s + diffHoras(t.creado, t.actualizado), 0);
        tiempoPromedio = Math.round((totalHoras / cerrados.length) * 10) / 10;
      }
      return {
        asesor: a,
        asignados: propios.length,
        gestionados: propios.filter((t) => t.estado !== "Abierto / Recibido").length,
        cerrados: cerrados.length,
        pendientes: propios.filter((t) => t.estado === "Abierto / Recibido").length,
        tiempoPromedio,
      };
    });
  }, [ticketsFiltrados]);

  // ----- Gestión por área -----
  const gestionAreas = useMemo(
    () => AREAS.map((a, i) => ({ area: a, cantidad: ticketsFiltrados.filter((t) => t.area === a).length, color: COLOR_AREA[i] })),
    [ticketsFiltrados]
  );

  // ----- Tickets diarios (creación) -----
  const ticketsDiarios = useMemo(() => {
    const map = {};
    ticketsFiltrados.forEach((t) => { const d = t.creado.slice(0, 10); map[d] = (map[d] || 0) + 1; });
    return Object.keys(map).sort().map((d) => ({ fecha: d.slice(5), cantidad: map[d] }));
  }, [ticketsFiltrados]);

  // ----- Gestión diaria (acciones registradas en el historial) -----
  const gestionDiaria = useMemo(() => {
    const map = {};
    ticketsFiltrados.forEach((t) => t.comentarios.forEach((c) => { const d = c.fecha.slice(0, 10); map[d] = (map[d] || 0) + 1; }));
    return Object.keys(map).sort().map((d) => ({ fecha: d.slice(5), cantidad: map[d] }));
  }, [ticketsFiltrados]);

  // ----- Distribución por categoría / prioridad -----
  const distCategorias = useMemo(
    () => categorias.map((c, i) => ({ nombre: c, valor: ticketsFiltrados.filter((t) => t.categoria === c).length, color: COLOR_CATEGORIA[i] })),
    [ticketsFiltrados]
  );
  const distPrioridad = useMemo(
    () => PRIORIDADES.map((p) => ({ nombre: p, valor: ticketsFiltrados.filter((t) => t.prioridad === p).length, color: PRIORIDAD_STYLE[p] })),
    [ticketsFiltrados]
  );

  // ----- Tendencias -----
  const tendencias = useMemo(() => {
    const lista = [];
    if (ticketsDiarios.length >= 2) {
      const mitad = Math.ceil(ticketsDiarios.length / 2);
      const primera = ticketsDiarios.slice(0, mitad).reduce((s, d) => s + d.cantidad, 0);
      const segunda = ticketsDiarios.slice(mitad).reduce((s, d) => s + d.cantidad, 0);
      if (segunda > primera) lista.push({ tipo: "up", texto: `Los tickets creados aumentaron en la segunda mitad del período (${primera} → ${segunda}).` });
      else if (segunda < primera) lista.push({ tipo: "down", texto: `Los tickets creados disminuyeron en la segunda mitad del período (${primera} → ${segunda}).` });
      else lista.push({ tipo: "flat", texto: "El volumen de tickets se mantiene estable en el período analizado." });
    }
    const areaTop = gestionAreas.slice().sort((a, b) => b.cantidad - a.cantidad)[0];
    if (areaTop && areaTop.cantidad > 0) lista.push({ tipo: "info", texto: `${areaTop.area} es el área con mayor cantidad de casos (${areaTop.cantidad}).` });
    const asesorTop = gestionAsesores.slice().sort((a, b) => b.asignados - a.asignados)[0];
    if (asesorTop && asesorTop.asignados > 0) lista.push({ tipo: "info", texto: `${asesorTop.asesor} tiene la mayor carga de tickets asignados (${asesorTop.asignados}).` });
    const catTop = distCategorias.slice().sort((a, b) => b.valor - a.valor)[0];
    if (catTop && catTop.valor > 0) lista.push({ tipo: "info", texto: `"${catTop.nombre}" es la categoría con mayor recurrencia (${catTop.valor} casos).` });
    return lista;
  }, [ticketsDiarios, gestionAreas, gestionAsesores, distCategorias]);

  // ----- Exportación (respeta los filtros aplicados) -----
  function construirFilas() {
    const rows = [[
      "Ticket", "Estudiante", "Identificacion", "Categoria", "Subcategoria", "Area", "Asesor",
      "Prioridad", "Estado", "Creado", "Descripcion del estudiante",
      "Respuestas de los asesores al estudiante", "Novedad de area / Remision",
    ]];
    ticketsFiltrados.forEach((t) => {
      const respuestasTexto = (t.historialRespuestas || [])
        .map((r) => `[${r.fecha}] ${r.texto}`)
        .join(" | ");
      rows.push([
        "TCK-" + String(t.id).padStart(4, "0"), t.nombre, t.codigo, t.categoria, t.subcategoria || "",
        t.area, t.asesor, t.prioridad, t.estado, t.creado,
        t.descripcion || "",
        respuestasTexto,
        t.novedadArea || "",
      ]);
    });
    return rows;
  }
  function exportarCSV() {
    const rows = construirFilas();
    const csv = rows.map((r) => r.map((c) => '"' + String(c).replace(/"/g, '""') + '"').join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "reporte_tickets.csv"; a.click();
  }
  function exportarExcel() {
    const rows = construirFilas();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [
      { wch: 10 }, { wch: 20 }, { wch: 16 }, { wch: 22 }, { wch: 30 }, { wch: 18 }, { wch: 16 },
      { wch: 10 }, { wch: 18 }, { wch: 17 }, { wch: 45 }, { wch: 55 }, { wch: 45 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tickets");
    XLSX.writeFile(wb, "reporte_tickets.xlsx");
  }
  // Arquitectura preparada para exportación a PDF en una futura iteración.
  function exportarPDF() {
    alert("La exportación a PDF estará disponible en una próxima versión.");
  }

  const inputStyle = { border: "1px solid #DDE2E9" };
  const selects = [
    { label: "ESTADO", valor: fEstado, set: setFEstado, opciones: ["Todos", ...ESTADOS] },
    { label: "ÁREA", valor: fArea, set: setFArea, opciones: ["Todas", ...AREAS] },
    { label: "ASESOR", valor: fAsesor, set: setFAsesor, opciones: ["Todos", ...asesoresReales] },
    { label: "CATEGORÍA", valor: fCategoria, set: setFCategoria, opciones: ["Todas", ...categorias] },
    { label: "PRIORIDAD", valor: fPrioridad, set: setFPrioridad, opciones: ["Todas", ...PRIORIDADES] },
  ];

  return (
    <div>
      {/* ---------- Filtros ---------- */}
      <div className="p-4 rounded-xl mb-5" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="flex items-center gap-1.5 text-xs font-bold mb-3" style={{ color: "#7A8698", letterSpacing: "0.05em" }}>
          <FilterIcon size={13} /> FILTROS
          {hayFiltrosActivos && (
            <button onClick={limpiarFiltros} className="ml-auto flex items-center gap-1 text-xs font-semibold" style={{ color: "#D64545" }}>
              <X size={12} /> Limpiar
            </button>
          )}
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3">
          <div>
            <label className="text-[10px] font-semibold block mb-1" style={{ color: "#9AA4B2" }}>DESDE</label>
            <input type="date" value={fInicio} onChange={(e) => setFInicio(e.target.value)}
              className="w-full px-2 py-1.5 rounded-lg text-xs" style={inputStyle} />
          </div>
          <div>
            <label className="text-[10px] font-semibold block mb-1" style={{ color: "#9AA4B2" }}>HASTA</label>
            <input type="date" value={fFin} onChange={(e) => setFFin(e.target.value)}
              className="w-full px-2 py-1.5 rounded-lg text-xs" style={inputStyle} />
          </div>
          {selects.map((s) => (
            <div key={s.label}>
              <label className="text-[10px] font-semibold block mb-1" style={{ color: "#9AA4B2" }}>{s.label}</label>
              <select value={s.valor} onChange={(e) => s.set(e.target.value)} className="w-full px-2 py-1.5 rounded-lg text-xs" style={inputStyle}>
                {s.opciones.map((o) => <option key={o}>{o}</option>)}
              </select>
            </div>
          ))}
        </div>
      </div>

      {/* ---------- KPIs ---------- */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4 mb-6">
        <TarjetaKPI titulo="Total de tickets" valor={kpis.total} porcentaje={100} variacion={variacion ? variacion.total : null} color="#1B2430" Icono={Inbox} />
        {ESTADOS.map((e) => (
          <TarjetaKPI key={e} titulo={e} valor={kpis.porEstado[e]}
            porcentaje={kpis.total ? Math.round((kpis.porEstado[e] / kpis.total) * 100) : 0}
            variacion={variacion ? variacion[e] : null}
            color={ESTADO_STYLE[e].dot} Icono={ESTADO_ICONO[e]} />
        ))}
      </div>

      {/* ---------- Gestión por asesor ---------- */}
      <div className="p-5 rounded-xl mb-6" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
          <Users size={15} color="#1652F0" /> Gestión por asesor
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <div style={{ width: "100%", height: 240 }}>
            <ResponsiveContainer>
              <BarChart data={gestionAsesores} layout="vertical" margin={{ left: 10, right: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11 }} allowDecimals={false} />
                <YAxis type="category" dataKey="asesor" tick={{ fontSize: 11 }} width={90} />
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="asignados" name="Asignados" fill="#1652F0" radius={[0, 4, 4, 0]} />
                <Bar dataKey="cerrados" name="Cerrados" fill="#2FAE6B" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr style={{ color: "#9AA4B2" }}>
                  <th className="text-left font-semibold pb-2">Asesor</th>
                  <th className="text-right font-semibold pb-2">Asignados</th>
                  <th className="text-right font-semibold pb-2">Gestionados</th>
                  <th className="text-right font-semibold pb-2">Cerrados</th>
                  <th className="text-right font-semibold pb-2">Pendientes</th>
                  <th className="text-right font-semibold pb-2">T. promedio</th>
                </tr>
              </thead>
              <tbody>
                {gestionAsesores.map((a) => (
                  <tr key={a.asesor} style={{ borderTop: "1px solid #F0F2F5" }}>
                    <td className="py-2 font-medium" style={{ color: "#1B2430" }}>{a.asesor}</td>
                    <td className="text-right">{a.asignados}</td>
                    <td className="text-right">{a.gestionados}</td>
                    <td className="text-right">{a.cerrados}</td>
                    <td className="text-right">{a.pendientes}</td>
                    <td className="text-right">{a.tiempoPromedio !== null ? `${a.tiempoPromedio} h` : "N/D"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* ---------- Gestión por área ---------- */}
      <div className="p-5 rounded-xl mb-6" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
          <Building2 size={15} color="#1652F0" /> Gestión por área
        </div>
        <div style={{ width: "100%", height: 260 }}>
          <ResponsiveContainer>
            <BarChart data={gestionAreas} margin={{ left: 0, right: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" vertical={false} />
              <XAxis dataKey="area" tick={{ fontSize: 10 }} interval={0} angle={-15} textAnchor="end" height={60} />
              <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
              <Tooltip />
              <Bar dataKey="cantidad" name="Tickets" radius={[6, 6, 0, 0]}>
                {gestionAreas.map((g, i) => <Cell key={i} fill={g.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* ---------- Tickets diarios / Gestión diaria ---------- */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 mb-6">
        <div className="p-5 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
          <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
            <Activity size={15} color="#1652F0" /> Tickets creados por día
          </div>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <LineChart data={ticketsDiarios}>
                <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" />
                <XAxis dataKey="fecha" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                <Tooltip />
                <Line type="monotone" dataKey="cantidad" name="Creados" stroke="#1652F0" strokeWidth={2} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="p-5 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
          <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
            <History size={15} color="#2FAE6B" /> Gestión diaria (acciones registradas)
          </div>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <LineChart data={gestionDiaria}>
                <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" />
                <XAxis dataKey="fecha" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                <Tooltip />
                <Line type="monotone" dataKey="cantidad" name="Gestionados" stroke="#2FAE6B" strokeWidth={2} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* ---------- Distribución por categoría / prioridad ---------- */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 mb-6">
        <div className="p-5 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
          <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
            <PieIcon size={15} color="#1652F0" /> Distribución por categoría
          </div>
          <div style={{ width: "100%", height: 240 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={distCategorias} dataKey="valor" nameKey="nombre" innerRadius={55} outerRadius={85} paddingAngle={2}>
                  {distCategorias.map((c, i) => <Cell key={i} fill={c.color} />)}
                </Pie>
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: 11 }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="p-5 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
          <div className="flex items-center gap-1.5 text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>
            <BarChart3 size={15} color="#1652F0" /> Distribución por prioridad
          </div>
          <div style={{ width: "100%", height: 240 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={distPrioridad} dataKey="valor" nameKey="nombre" innerRadius={55} outerRadius={85} paddingAngle={2}>
                  {distPrioridad.map((p, i) => <Cell key={i} fill={p.color} />)}
                </Pie>
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: 11 }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* ---------- Tendencias ---------- */}
      <div className="p-5 rounded-xl mb-6" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="text-sm font-semibold mb-4" style={{ color: "#1B2430" }}>Tendencias de la operación</div>
        {tendencias.length === 0 && <div className="text-xs" style={{ color: "#B7BFCB" }}>No hay suficiente información para calcular tendencias con los filtros actuales.</div>}
        <div className="space-y-2">
          {tendencias.map((t, i) => {
            const Icono = t.tipo === "up" ? TrendingUp : t.tipo === "down" ? TrendingDown : t.tipo === "flat" ? Minus : Activity;
            const color = t.tipo === "up" ? "#2FAE6B" : t.tipo === "down" ? "#D64545" : "#1652F0";
            return (
              <div key={i} className="flex items-start gap-2 text-sm p-2.5 rounded-lg" style={{ background: "#F5F7FB" }}>
                <Icono size={15} color={color} className="mt-0.5" />
                <span style={{ color: "#1B2430" }}>{t.texto}</span>
              </div>
            );
          })}
        </div>
      </div>

      {/* ---------- Exportación ---------- */}
      <div className="p-5 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="text-sm font-semibold mb-1" style={{ color: "#1B2430" }}>Exportar reporte</div>
        <div className="text-xs mb-3" style={{ color: "#9AA4B2" }}>La exportación respeta los filtros aplicados ({ticketsFiltrados.length} tickets).</div>
        <div className="flex flex-wrap gap-2">
          <button onClick={exportarExcel} className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1B2430", color: "#FFFFFF" }}>
            <FileSpreadsheet size={14} /> Excel (.xlsx)
          </button>
          <button onClick={exportarCSV} className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#FFFFFF", color: "#1B2430", border: "1px solid #DDE2E9" }}>
            <Download size={14} /> CSV
          </button>
          <button onClick={exportarPDF} className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#F4F6F9", color: "#9AA4B2" }}>
            <FileText size={14} /> PDF (próximamente)
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Módulo: Gestión de asesores (exclusivo del Líder de Área) ----------
function PanelAsesores({ agentes, onCambio, agenteActual }) {
  const [creando, setCreando] = useState(false);
  const [editandoId, setEditandoId] = useState(null);
  const [form, setForm] = useState({ nombre: "", usuario: "", password: "" });
  const [error, setError] = useState("");
  const [guardando, setGuardando] = useState(false);

  const asesores = agentes.filter((a) => a.rol === ROL_ASESOR);

  function iniciarCreacion() {
    setCreando(true);
    setEditandoId(null);
    setForm({ nombre: "", usuario: "", password: "" });
    setError("");
  }

  function iniciarEdicion(a) {
    setEditandoId(a.id);
    setCreando(false);
    setForm({ nombre: a.nombre, usuario: a.usuario, password: "" });
    setError("");
  }

  function cancelar() {
    setCreando(false);
    setEditandoId(null);
    setError("");
  }

  async function guardarNuevo() {
    if (!form.nombre.trim() || !form.usuario.trim() || !form.password.trim()) {
      setError("Completa nombre, usuario y contraseña.");
      return;
    }
    setGuardando(true);
    setError("");
    try {
      await apiPost("/api/agentes", { nombre: form.nombre.trim(), usuario: form.usuario.trim(), password: form.password.trim() });
      await onCambio();
      setCreando(false);
    } catch (err) {
      setError(err.message);
    } finally {
      setGuardando(false);
    }
  }

  async function guardarEdicion(id) {
    if (!form.nombre.trim() || !form.usuario.trim()) {
      setError("El nombre y el usuario no pueden estar vacíos.");
      return;
    }
    setGuardando(true);
    setError("");
    try {
      const payload = { nombre: form.nombre.trim(), usuario: form.usuario.trim() };
      if (form.password.trim()) payload.password = form.password.trim();
      await apiPatch(`/api/agentes/${id}`, payload);
      await onCambio();
      setEditandoId(null);
    } catch (err) {
      setError(err.message);
    } finally {
      setGuardando(false);
    }
  }

  async function alternarActivo(a) {
    try {
      await apiPatch(`/api/agentes/${a.id}`, { activo: !a.activo });
      await onCambio();
    } catch (err) {
      setError(err.message);
    }
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <div>
          <div className="text-sm font-semibold" style={{ color: "#1B2430" }}>Cuentas de asesores</div>
          <div className="text-xs" style={{ color: "#9AA4B2" }}>Solo el Líder de Área puede crear, editar o desactivar cuentas.</div>
        </div>
        {!creando && (
          <button onClick={iniciarCreacion} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1652F0", color: "#FFFFFF" }}>
            + Nuevo asesor
          </button>
        )}
      </div>

      {creando && (
        <div className="p-4 rounded-xl mb-4" style={{ background: "#FFFFFF", border: "1px solid #1652F0" }}>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-3">
            <div>
              <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>NOMBRE COMPLETO</label>
              <input value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
            </div>
            <div>
              <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>USUARIO</label>
              <input value={form.usuario} onChange={(e) => setForm({ ...form, usuario: e.target.value })}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
            </div>
            <div>
              <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>CONTRASEÑA</label>
              <input value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
            </div>
          </div>
          {error && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{error}</div>}
          <div className="flex gap-2">
            <button onClick={guardarNuevo} disabled={guardando} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1B2430", color: "#FFFFFF" }}>{guardando ? "Guardando..." : "Guardar"}</button>
            <button onClick={cancelar} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#F4F6F9", color: "#5A6577" }}>Cancelar</button>
          </div>
        </div>
      )}

      <div className="rounded-xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        {asesores.map((a) => (
          <div key={a.id} className="p-4" style={{ borderBottom: "1px solid #F0F2F5" }}>
            {editandoId === a.id ? (
              <div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-3">
                  <div>
                    <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>NOMBRE</label>
                    <input value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })}
                      className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                  </div>
                  <div>
                    <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>USUARIO</label>
                    <input value={form.usuario} onChange={(e) => setForm({ ...form, usuario: e.target.value })}
                      className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                  </div>
                  <div>
                    <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>NUEVA CONTRASEÑA</label>
                    <input value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })}
                      placeholder="Dejar vacío para no cambiarla"
                      className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                  </div>
                </div>
                {error && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{error}</div>}
                <div className="flex gap-2">
                  <button onClick={() => guardarEdicion(a.id)} disabled={guardando} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1B2430", color: "#FFFFFF" }}>{guardando ? "Guardando..." : "Guardar cambios"}</button>
                  <button onClick={cancelar} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#F4F6F9", color: "#5A6577" }}>Cancelar</button>
                </div>
              </div>
            ) : (
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div className="flex items-center gap-3">
                  <div style={{ width: 36, height: 36, borderRadius: 999, background: "#F4F6F9", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>
                    {a.avatar || "🙂"}
                  </div>
                  <div>
                    <div className="text-sm font-semibold" style={{ color: "#1B2430" }}>{a.nombre}</div>
                    <div className="text-xs" style={{ color: "#9AA4B2" }}>@{a.usuario} · Creado {a.fechaCreacion}</div>
                  </div>
                  <span className="text-xs font-semibold px-2.5 py-1 rounded-full" style={{ background: a.activo ? "#E6F6ED" : "#FCEAEA", color: a.activo ? "#1D7A47" : "#A62E2E" }}>
                    {a.activo ? "Activo" : "Desactivado"}
                  </span>
                </div>
                <div className="flex gap-2">
                  <button onClick={() => iniciarEdicion(a)} className="px-3 py-1.5 rounded-lg text-xs font-semibold" style={{ background: "#F4F6F9", color: "#5A6577" }}>Editar</button>
                  <button onClick={() => alternarActivo(a)}
                    className="px-3 py-1.5 rounded-lg text-xs font-semibold"
                    style={{ background: a.activo ? "#FCEAEA" : "#E6F6ED", color: a.activo ? "#A62E2E" : "#1D7A47" }}>
                    {a.activo ? "Desactivar" : "Activar"}
                  </button>
                </div>
              </div>
            )}
          </div>
        ))}
        {asesores.length === 0 && (
          <div className="p-6 text-center text-sm" style={{ color: "#9AA4B2" }}>Aún no hay asesores registrados.</div>
        )}
      </div>
    </div>
  );
}

// ---------- Módulo: Mi Perfil ----------
function PanelMiPerfil({ agenteActual, onCambio }) {
  const [guardando, setGuardando] = useState(false);
  const [error, setError] = useState("");
  const esLider = agenteActual.rol === ROL_LIDER;

  const [editandoDatos, setEditandoDatos] = useState(false);
  const [formNombre, setFormNombre] = useState(agenteActual.nombre);
  const [formUsuario, setFormUsuario] = useState(agenteActual.usuario);

  const [passwordActual, setPasswordActual] = useState("");
  const [passwordNueva, setPasswordNueva] = useState("");
  const [passwordConfirmar, setPasswordConfirmar] = useState("");
  const [errorPassword, setErrorPassword] = useState("");
  const [exitoPassword, setExitoPassword] = useState(false);
  const [guardandoPassword, setGuardandoPassword] = useState(false);

  async function guardarPassword() {
    setErrorPassword("");
    setExitoPassword(false);
    if (!passwordActual.trim() || !passwordNueva.trim()) {
      setErrorPassword("Completa tu contraseña actual y la nueva.");
      return;
    }
    if (passwordNueva.trim().length < 4) {
      setErrorPassword("La nueva contraseña debe tener al menos 4 caracteres.");
      return;
    }
    if (passwordNueva.trim() !== passwordConfirmar.trim()) {
      setErrorPassword("Las contraseñas nuevas no coinciden.");
      return;
    }
    setGuardandoPassword(true);
    try {
      await apiPatch(`/api/agentes/${agenteActual.id}`, {
        passwordActual: passwordActual.trim(),
        password: passwordNueva.trim(),
      });
      setPasswordActual("");
      setPasswordNueva("");
      setPasswordConfirmar("");
      setExitoPassword(true);
      setTimeout(() => setExitoPassword(false), 2500);
    } catch (err) {
      setErrorPassword(err.message);
    } finally {
      setGuardandoPassword(false);
    }
  }

  async function actualizarAvatar(valor) {
    setGuardando(true);
    setError("");
    try {
      await apiPatch(`/api/agentes/${agenteActual.id}`, { avatar: valor });
      await onCambio();
    } catch (err) {
      setError(err.message);
    } finally {
      setGuardando(false);
    }
  }

  function iniciarEdicionDatos() {
    setFormNombre(agenteActual.nombre);
    setFormUsuario(agenteActual.usuario);
    setError("");
    setEditandoDatos(true);
  }

  async function guardarDatos() {
    if (!formNombre.trim() || !formUsuario.trim()) {
      setError("El nombre y el usuario no pueden estar vacíos.");
      return;
    }
    setGuardando(true);
    setError("");
    try {
      await apiPatch(`/api/agentes/${agenteActual.id}`, { nombre: formNombre.trim(), usuario: formUsuario.trim() });
      await onCambio();
      setEditandoDatos(false);
    } catch (err) {
      setError(err.message);
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="max-w-2xl">
      <div className="p-6 rounded-xl mb-5" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="flex items-center gap-4 mb-5">
          <div style={{ width: 64, height: 64, borderRadius: 999, background: "#F4F6F9", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 30, overflow: "hidden" }}>
            {agenteActual.avatar || "🙂"}
          </div>
          <div>
            <div className="text-lg font-semibold" style={{ color: "#1B2430" }}>{agenteActual.nombre}</div>
            <div className="text-xs" style={{ color: "#9AA4B2" }}>@{agenteActual.usuario} · {agenteActual.rol}</div>
          </div>
        </div>

        {editandoDatos ? (
          <div className="mb-4">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <div>
                <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#9AA4B2" }}>NOMBRE COMPLETO</label>
                <input value={formNombre} onChange={(e) => setFormNombre(e.target.value)}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              </div>
              <div>
                <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#9AA4B2" }}>USUARIO</label>
                <input value={formUsuario} onChange={(e) => setFormUsuario(e.target.value)}
                  className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              </div>
            </div>
            {error && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{error}</div>}
            <div className="flex gap-2">
              <button onClick={guardarDatos} disabled={guardando} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1652F0", color: "#FFFFFF" }}>
                {guardando ? "Guardando..." : "Guardar cambios"}
              </button>
              <button onClick={() => setEditandoDatos(false)} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#F4F6F9", color: "#5A6577" }}>
                Cancelar
              </button>
            </div>
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-5 text-sm">
            <div>
              <div className="text-[11px] font-semibold" style={{ color: "#9AA4B2" }}>NOMBRE COMPLETO</div>
              <div style={{ color: "#1B2430" }}>{agenteActual.nombre}</div>
            </div>
            <div>
              <div className="text-[11px] font-semibold" style={{ color: "#9AA4B2" }}>ROL</div>
              <div style={{ color: "#1B2430" }}>{agenteActual.rol}</div>
            </div>
            <div>
              <div className="text-[11px] font-semibold" style={{ color: "#9AA4B2" }}>CUENTA CREADA</div>
              <div style={{ color: "#1B2430" }}>{agenteActual.fechaCreacion || "—"}</div>
            </div>
          </div>
        )}

        {esLider ? (
          !editandoDatos && (
            <button onClick={iniciarEdicionDatos} className="text-xs font-semibold" style={{ color: "#1652F0" }}>
              Editar nombre y usuario
            </button>
          )
        ) : (
          <div className="text-xs" style={{ color: "#B7BFCB" }}>
            El nombre, usuario, rol y permisos solo pueden ser modificados por el Líder de Área.
          </div>
        )}
      </div>

      <div className="p-6 rounded-xl mb-5" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
        <div className="text-sm font-semibold mb-3" style={{ color: "#1B2430" }}>Avatar</div>
        <div className="text-xs mb-2" style={{ color: "#9AA4B2" }}>Elige un avatar predeterminado:</div>
        {error && <div className="text-xs mb-2" style={{ color: "#D64545" }}>{error}</div>}
        <div className="flex flex-wrap gap-2">
          {AVATARES.map((av) => (
            <button key={av} onClick={() => actualizarAvatar(av)} disabled={guardando}
              className="flex items-center justify-center"
              style={{
                width: 42, height: 42, borderRadius: 999, fontSize: 20,
                background: agenteActual.avatar === av ? "#EEF3FF" : "#F4F6F9",
                border: agenteActual.avatar === av ? "2px solid #1652F0" : "2px solid transparent",
              }}>
              {av}
            </button>
          ))}
        </div>
      </div>

      <div className="p-6 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC", opacity: esLider ? 1 : 0.6 }}>
        <div className="text-sm font-semibold mb-1" style={{ color: "#1B2430" }}>Cambiar contraseña</div>
        {esLider ? (
          <>
            <div className="text-xs mb-3" style={{ color: "#9AA4B2" }}>Debes confirmar tu contraseña actual para poder cambiarla.</div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-3">
              <input type="password" placeholder="Contraseña actual" value={passwordActual} onChange={(e) => setPasswordActual(e.target.value)}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              <input type="password" placeholder="Nueva contraseña" value={passwordNueva} onChange={(e) => setPasswordNueva(e.target.value)}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
              <input type="password" placeholder="Confirmar nueva contraseña" value={passwordConfirmar} onChange={(e) => setPasswordConfirmar(e.target.value)}
                className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
            </div>
            {errorPassword && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{errorPassword}</div>}
            <button onClick={guardarPassword} disabled={guardandoPassword}
              className="px-4 py-2 rounded-lg text-sm font-semibold"
              style={{ background: exitoPassword ? "#2FAE6B" : "#1652F0", color: "#FFFFFF" }}>
              {guardandoPassword ? "Guardando..." : exitoPassword ? "Contraseña actualizada" : "Actualizar contraseña"}
            </button>
          </>
        ) : (
          <>
            <div className="text-xs mb-3" style={{ color: "#9AA4B2" }}>Disponible próximamente.</div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <input disabled placeholder="Contraseña actual" className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9", background: "#F5F7FB" }} />
              <input disabled placeholder="Nueva contraseña" className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9", background: "#F5F7FB" }} />
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ---------- App principal ----------
export default function App() {
  const [vista, setVista] = useState("home"); // 'home' | 'estudiante' | 'soporte' | 'soporte-dashboard'
  const [agenteActual, setAgenteActual] = useState(null);
  const [verificandoSesion, setVerificandoSesion] = useState(true);

  // Al cargar la página, preguntamos a la API si ya hay una sesión activa (cookie),
  // para que el asesor no tenga que volver a iniciar sesión cada vez que recarga.
  useEffect(() => {
    (async () => {
      try {
        const datos = await apiGet("/api/auth/me");
        setAgenteActual(datos);
        setVista("soporte-dashboard");
      } catch {
        // No hay sesión activa, se queda en la portada.
      } finally {
        setVerificandoSesion(false);
      }
    })();
  }, []);

  async function handleLogout() {
    try { await apiPost("/api/auth/logout", {}); } catch { /* seguimos igual */ }
    setAgenteActual(null);
    setVista("home");
  }

  if (verificandoSesion) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: "#F5F7FB", color: "#5B6472" }}>
        Cargando...
      </div>
    );
  }

  return (
    <div style={{ fontFamily: "Inter, ui-sans-serif, system-ui, sans-serif" }}>
      <EstiloGlobalInstitucional />
      {vista === "home" && <Portada onSelect={setVista} />}
      {vista === "estudiante" && (
        <VistaEstudiante onBack={() => setVista("home")} />
      )}
      {vista === "soporte" && (
        <LoginSoporte onLogin={(datos) => { setAgenteActual(datos); setVista("soporte-dashboard"); }} onBack={() => setVista("home")} />
      )}
      {vista === "soporte-dashboard" && agenteActual && (
        <DashboardSoporte agenteActual={agenteActual} onLogout={handleLogout} />
      )}
    </div>
  );
}
