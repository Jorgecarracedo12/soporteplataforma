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

// ---------- Datos de referencia ----------
const CATEGORIAS = ["Plataforma", "Acceso", "Virtual", "Otro"];

// Áreas internas a las que se puede remitir un caso (uso exclusivo del panel de soporte)
const AREAS = ["Prizma", "Partikle", "Problemas LDAP", "Coursera", "Cargue", "Coordinación Académica", "Admisiones"];
const COLOR_AREA = ["#2F6FED", "#0BA5EC", "#7C5CFC", "#F59E0B", "#EF4444", "#10B981", "#64748B"];
const COLOR_CATEGORIA = ["#2F6FED", "#7C5CFC", "#94A3B8"];

const PRIORIDADES = ["Alta", "Media", "Baja"];
const PRIORIDAD_STYLE = { "Alta": "#D64545", "Media": "#E8A93B", "Baja": "#8FA3BF" };

// Estados del ciclo de vida del ticket, gestionados por el equipo de soporte
const ESTADOS = ["Abierto / Recibido", "En gestión", "Novedad", "Cerrado / Subsanado"];
const ESTADO_ICONO = { "Abierto / Recibido": Clock, "En gestión": Settings2, "Novedad": AlertTriangle, "Cerrado / Subsanado": CheckCircle2 };

// Asesores disponibles para asignar un ticket (se derivan de los usuarios del panel de soporte)
const ASESORES = ["Sin asignar", ...["Ana Torres", "Luis Ramírez"]];

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
  "En gestión": { bg: "#E9F0FE", fg: "#1D4FB8", dot: "#2F6FED" },
  "Novedad": { bg: "#FCEAEA", fg: "#A62E2E", dot: "#D64545" },
  "Cerrado / Subsanado": { bg: "#E6F6ED", fg: "#1D7A47", dot: "#2FAE6B" },
};

const AGENTES = [
  { usuario: "soporte1", password: "1234", nombre: "Ana Torres" },
  { usuario: "soporte2", password: "1234", nombre: "Luis Ramírez" },
];

// ---------- Datos semilla (demo) ----------
const SEED_TICKETS = [
  {
    id: 1,
    nombre: "Camila Restrepo",
    codigo: "A0123456",
    correo: "camila.restrepo@correo.edu",
    prizmaUsuario: "camila.restrepo",
    prizmaPassword: "cr_2026*",
    categoria: "Plataforma Prizma",
    subcategoria: "No recuerdo mis credenciales",
    descripcion: "No puedo cargar el material del módulo 3, la página se queda cargando.",
    evidencia: null,
    area: "Prizma",
    prioridad: "Alta",
    estado: "Abierto / Recibido",
    asesor: "Sin asignar",
    respuestaEstudiante: "",
    novedadArea: "",
    creado: "2026-07-27 09:12",
    actualizado: "2026-07-27 09:12",
    comentarios: [],
  },
  {
    id: 2,
    nombre: "Julián Pérez",
    codigo: "A0129981",
    correo: "julian.perez@correo.edu",
    prizmaUsuario: "julian.perez",
    prizmaPassword: "jp_clave1",
    categoria: "Plataforma Prizma",
    subcategoria: "No sé usar la plataforma",
    descripcion: "Olvidé mis credenciales y el link de recuperación no me llega al correo.",
    evidencia: null,
    area: "Partikle",
    prioridad: "Media",
    estado: "En gestión",
    asesor: "Ana Torres",
    respuestaEstudiante: "Estamos verificando tu usuario en la plataforma, te contactaremos en las próximas horas.",
    novedadArea: "Se solicitó a Partikle validar el estado de la cuenta del estudiante.",
    creado: "2026-07-26 15:40",
    actualizado: "2026-07-27 08:05",
    comentarios: [
      { autor: "Ana Torres", texto: "Se remitió el caso a Partikle para validar la cuenta.", fecha: "2026-07-27 08:05" },
    ],
  },
  {
    id: 3,
    nombre: "Mariana Gómez",
    codigo: "A0110044",
    correo: "mariana.gomez@correo.edu",
    prizmaUsuario: "mariana.gomez",
    prizmaPassword: "mg_2026*",
    categoria: "Inconveniente con Asignaturas",
    subcategoria: "No visualizo ninguna de mis asignaturas.",
    descripcion: "No veo ninguna asignatura matriculada en la plataforma desde ayer.",
    evidencia: null,
    area: "Coordinación Académica",
    prioridad: "Alta",
    estado: "Cerrado / Subsanado",
    asesor: "Luis Ramírez",
    respuestaEstudiante: "Tu caso fue resuelto: ya puedes visualizar todas tus asignaturas matriculadas.",
    novedadArea: "Se sincronizó la matrícula con Coordinación Académica.",
    creado: "2026-07-25 10:02",
    actualizado: "2026-07-25 12:30",
    comentarios: [
      { autor: "Luis Ramírez", texto: "Se sincronizó la matrícula y se confirmó con el estudiante.", fecha: "2026-07-25 12:30" },
    ],
  },
];

function nextTicketNum(list) {
  const max = list.reduce((m, t) => Math.max(m, t.id), 0);
  return "TCK-" + String(max + 1).padStart(4, "0");
}

function nowStr() {
  const d = new Date();
  return d.toISOString().slice(0, 16).replace("T", " ");
}

function diffHoras(fechaInicioStr, fechaFinStr) {
  const a = new Date(fechaInicioStr.replace(" ", "T"));
  const b = new Date(fechaFinStr.replace(" ", "T"));
  return Math.max(0, (b - a) / (1000 * 60 * 60));
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
        color: "#F7F8FA",
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

// ---------- Vista: Portada ----------
function Portada({ onSelect }) {
  return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: "#F7F8FA" }}>
      <div className="w-full max-w-3xl px-6">
        <div className="text-center mb-10">
          <div
            className="inline-block mb-4"
            style={{
              fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
              fontSize: 12,
              letterSpacing: "0.15em",
              color: "#7A8698",
              textTransform: "uppercase",
            }}
          >
            Mesa de ayuda · Soporte estudiantil
          </div>
          <h1 className="text-4xl font-bold" style={{ color: "#1B2430", letterSpacing: "-0.02em" }}>
            Centro de Tickets
          </h1>
          <p className="mt-3 text-base" style={{ color: "#5A6577" }}>
            Reporta un problema técnico o gestiona los tickets del equipo de soporte.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          <button
            onClick={() => onSelect("estudiante")}
            className="text-left p-6 rounded-xl transition"
            style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}
          >
            <div className="flex items-center justify-between mb-3">
              <span style={{ fontSize: 12, fontWeight: 700, color: "#2F6FED", letterSpacing: "0.06em" }}>
                ESTUDIANTE
              </span>
              <span style={{ color: "#B7BFCB" }}>→</span>
            </div>
            <div className="text-lg font-semibold mb-1" style={{ color: "#1B2430" }}>
              Reportar un problema
            </div>
            <div className="text-sm" style={{ color: "#5A6577" }}>
              Crea un ticket y consulta el estado de tus solicitudes anteriores.
            </div>
          </button>

          <button
            onClick={() => onSelect("soporte")}
            className="text-left p-6 rounded-xl transition"
            style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}
          >
            <div className="flex items-center justify-between mb-3">
              <span style={{ fontSize: 12, fontWeight: 700, color: "#1B2430", letterSpacing: "0.06em" }}>
                EQUIPO DE SOPORTE
              </span>
              <span style={{ color: "#B7BFCB" }}>→</span>
            </div>
            <div className="text-lg font-semibold mb-1" style={{ color: "#1B2430" }}>
              Ingresar al panel
            </div>
            <div className="text-sm" style={{ color: "#5A6577" }}>
              Gestiona, redirige y da solución a los tickets recibidos.
            </div>
          </button>
        </div>

        <div className="mt-8 text-center text-xs" style={{ color: "#9AA4B2" }}>
          Prototipo de demostración — los datos se reinician al recargar la página.
        </div>
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
  evidencia: null, // { nombre, dataUrl } o null (opcional)
};

// ---------- Componente: input de evidencia con vista previa ----------
function CampoEvidencia({ evidencia, onChange }) {
  function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) return;
    const reader = new FileReader();
    reader.onload = () => onChange({ nombre: file.name, dataUrl: reader.result });
    reader.readAsDataURL(file);
  }

  return (
    <div className="mb-4">
      <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>
        ADJUNTAR EVIDENCIA (opcional)
      </label>
      {!evidencia && (
        <label
          className="flex items-center justify-center gap-2 w-full py-4 rounded-lg text-xs cursor-pointer"
          style={{ border: "1px dashed #C7CEDA", color: "#7A8698", background: "#FAFBFC" }}
        >
          Haz clic para subir una imagen (captura de pantalla, foto del error, etc.)
          <input type="file" accept="image/*" onChange={handleFile} className="hidden" />
        </label>
      )}
      {evidencia && (
        <div className="flex items-center gap-3 p-3 rounded-lg" style={{ border: "1px solid #E2E6EC" }}>
          <img src={evidencia.dataUrl} alt="Vista previa" className="rounded-md object-cover"
            style={{ width: 56, height: 56 }} />
          <div className="flex-1 text-xs truncate" style={{ color: "#5A6577" }}>{evidencia.nombre}</div>
          <button type="button" onClick={() => onChange(null)} className="text-xs font-semibold" style={{ color: "#D64545" }}>
            Quitar
          </button>
        </div>
      )}
    </div>
  );
}

// ---------- Vista: Estudiante ----------
function VistaEstudiante({ tickets, setTickets, onBack }) {
  const [modo, setModo] = useState("crear"); // 'crear' | 'consultar'
  const [form, setForm] = useState(FORM_VACIO);
  const [enviado, setEnviado] = useState(null);
  const [buscarCodigo, setBuscarCodigo] = useState("");
  const [resultados, setResultados] = useState(null);
  const [errorForm, setErrorForm] = useState("");

  const subcategoriasDisponibles = CATEGORIAS_ESTUDIANTE[form.categoria] || [];
  const requiereSubcategoria = subcategoriasDisponibles.length > 0;

  // Ticket activo del estudiante que está escribiendo su número de identificación (bloquea nuevo ticket)
  const ticketActivoExistente = useMemo(() => {
    const id = form.codigo.trim().toLowerCase();
    if (!id) return null;
    return tickets.find(
      (t) => t.codigo.trim().toLowerCase() === id && ESTADOS_BLOQUEAN_NUEVO_TICKET.includes(t.estado)
    ) || null;
  }, [tickets, form.codigo]);

  function handleCategoriaChange(categoria) {
    const subs = CATEGORIAS_ESTUDIANTE[categoria] || [];
    setForm({ ...form, categoria, subcategoria: subs[0] || "" });
  }

  function handleSubmit() {
    // 1) Validar que no tenga ya un ticket activo
    const id = form.codigo.trim().toLowerCase();
    const activo = tickets.find(
      (t) => t.codigo.trim().toLowerCase() === id && ESTADOS_BLOQUEAN_NUEVO_TICKET.includes(t.estado)
    );
    if (activo) {
      setErrorForm(
        `Ya tienes un ticket activo (TCK-${String(activo.id).padStart(4, "0")}, estado "${activo.estado}"). Debes esperar a que sea gestionado o cerrado antes de crear uno nuevo.`
      );
      return;
    }

    // 2) Validar campos obligatorios (todos menos evidencia)
    const faltaSubcategoria = requiereSubcategoria && !form.subcategoria;
    if (!form.nombre || !form.codigo || !form.correo || !form.prizmaUsuario || !form.prizmaPassword || !form.descripcion || faltaSubcategoria) {
      setErrorForm("Por favor completa todos los campos obligatorios (*).");
      return;
    }

    setErrorForm("");
    const nuevoId = tickets.reduce((m, t) => Math.max(m, t.id), 0) + 1;
    const nuevo = {
      id: nuevoId,
      ...form,
      area: "Sin asignar",
      prioridad: "Media",
      estado: "Abierto / Recibido",
      asesor: "Sin asignar",
      respuestaEstudiante: "",
      novedadArea: "",
      creado: nowStr(),
      actualizado: nowStr(),
      comentarios: [],
    };
    setTickets([nuevo, ...tickets]);
    setEnviado(nuevo);
  }

  function handleBuscar() {
    const r = tickets.filter(
      (t) => t.codigo.toLowerCase() === buscarCodigo.trim().toLowerCase()
    );
    setResultados(r);
  }

  return (
    <div className="min-h-screen" style={{ background: "#F7F8FA" }}>
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

            <CampoEvidencia evidencia={form.evidencia} onChange={(ev) => setForm({ ...form, evidencia: ev })} />

            {errorForm && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{errorForm}</div>}

            <button type="button" onClick={handleSubmit} disabled={!!ticketActivoExistente}
              className="w-full py-3 rounded-lg text-sm font-semibold"
              style={{
                background: ticketActivoExistente ? "#C7CEDA" : "#2F6FED",
                color: "#FFFFFF",
                cursor: ticketActivoExistente ? "not-allowed" : "pointer",
              }}>
              Enviar ticket
            </button>
          </div>
        )}

        {modo === "crear" && enviado && (
          <div className="p-6 rounded-xl text-center" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
            <div className="text-sm mb-3" style={{ color: "#5A6577" }}>Tu ticket fue registrado</div>
            <div className="flex justify-center mb-4"><TicketStub id={enviado.id} /></div>
            <p className="text-sm mb-1" style={{ color: "#1B2430" }}>Guarda este número para hacer seguimiento.</p>
            <p className="text-sm" style={{ color: "#5A6577" }}>El equipo de soporte revisará tu caso pronto.</p>
            <button onClick={() => { setEnviado(null); setForm(FORM_VACIO); }}
              className="mt-6 text-sm font-semibold" style={{ color: "#2F6FED" }}>
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
              <button type="button" onClick={handleBuscar} className="px-4 py-2 rounded-lg text-sm font-semibold" style={{ background: "#1B2430", color: "#FFFFFF" }}>
                Buscar
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

        <div className="text-xs font-semibold mb-1" style={{ color: "#2F6FED", letterSpacing: "0.03em" }}>
          {ticket.categoria}{ticket.subcategoria ? " · " + ticket.subcategoria : ""}
        </div>

        <div className="text-sm mb-3" style={{ color: "#1B2430" }}>{ticket.descripcion}</div>

        {ticket.evidencia && (
          <img src={ticket.evidencia.dataUrl} alt="Evidencia adjunta" className="rounded-lg mb-3"
            style={{ maxHeight: 120, objectFit: "cover" }} />
        )}

        {ticket.respuestaEstudiante && (
          <div className="mt-1 mb-3 p-3 rounded-lg text-sm" style={{ background: "#EEF3FF", color: "#1D4FB8" }}>
            <div className="text-xs font-semibold mb-1" style={{ color: "#2F6FED" }}>Respuesta del equipo de soporte</div>
            {ticket.respuestaEstudiante}
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

  function handleSubmit() {
    const agente = AGENTES.find((a) => a.usuario === usuario.trim() && a.password === password.trim());
    if (agente) onLogin(agente);
    else setError("Usuario o contraseña incorrectos.");
  }

  return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: "#F7F8FA" }}>
      <div className="w-full max-w-sm px-6">
        <button onClick={onBack} className="text-sm mb-6" style={{ color: "#5A6577" }}>← Volver</button>
        <div className="p-6 rounded-xl" style={{ background: "#FFFFFF", border: "1px solid #E2E6EC" }}>
          <h2 className="text-lg font-semibold mb-1" style={{ color: "#1B2430" }}>Panel de soporte</h2>
          <p className="text-sm mb-5" style={{ color: "#5A6577" }}>Ingresa con tu usuario asignado.</p>
          <div className="mb-3">
            <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>USUARIO</label>
            <input value={usuario} onChange={(e) => setUsuario(e.target.value)}
              className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} placeholder="soporte1" />
          </div>
          <div className="mb-2">
            <label className="block text-xs font-semibold mb-1" style={{ color: "#5A6577" }}>CONTRASEÑA</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") handleSubmit(); }}
              className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} placeholder="1234" />
          </div>
          {error && <div className="text-xs mb-3" style={{ color: "#D64545" }}>{error}</div>}
          <button type="button" onClick={handleSubmit} className="w-full py-3 rounded-lg text-sm font-semibold mt-3" style={{ background: "#2F6FED", color: "#FFFFFF" }}>
            Ingresar
          </button>
          <div className="text-xs mt-4 text-center" style={{ color: "#9AA4B2" }}>
            Demo: soporte1 / 1234
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Vista: Dashboard de soporte ----------
function DashboardSoporte({ agente, tickets, setTickets, onLogout }) {
  const [tab, setTab] = useState("tickets"); // 'tickets' | 'reportes'
  const [busqueda, setBusqueda] = useState("");
  const [filtroEstado, setFiltroEstado] = useState("Todos");
  const [filtroArea, setFiltroArea] = useState("Todas");
  const [filtroPrioridad, setFiltroPrioridad] = useState("Todas");
  const [seleccionadoId, setSeleccionadoId] = useState(null);
  const [mostrarPassword, setMostrarPassword] = useState(false);
  const [respuestaLocal, setRespuestaLocal] = useState("");
  const [novedadLocal, setNovedadLocal] = useState("");
  const [guardadoOk, setGuardadoOk] = useState(false);

  const filtrados = useMemo(() => {
    const q = busqueda.trim().toLowerCase();
    return tickets.filter((t) => {
      if (filtroEstado !== "Todos" && t.estado !== filtroEstado) return false;
      if (filtroArea !== "Todas" && t.area !== filtroArea) return false;
      if (filtroPrioridad !== "Todas" && t.prioridad !== filtroPrioridad) return false;
      if (q && !(t.nombre.toLowerCase().includes(q) || t.codigo.toLowerCase().includes(q))) return false;
      return true;
    });
  }, [tickets, filtroEstado, filtroArea, filtroPrioridad, busqueda]);

  const seleccionado = useMemo(
    () => tickets.find((t) => t.id === seleccionadoId) || null,
    [tickets, seleccionadoId]
  );

  // Al cambiar de ticket seleccionado, sincronizar los editores de texto locales
  useEffect(() => {
    setRespuestaLocal(seleccionado ? seleccionado.respuestaEstudiante || "" : "");
    setNovedadLocal(seleccionado ? seleccionado.novedadArea || "" : "");
    setMostrarPassword(false);
    setGuardadoOk(false);
  }, [seleccionadoId]);

  function registrarHistorial(id, texto) {
    setTickets((prev) =>
      prev.map((t) =>
        t.id === id
          ? { ...t, comentarios: [...t.comentarios, { autor: agente.nombre, texto, fecha: nowStr() }] }
          : t
      )
    );
  }

  function actualizarCampo(id, campo, valor, etiqueta) {
    setTickets((prev) =>
      prev.map((t) => (t.id === id ? { ...t, [campo]: valor, actualizado: nowStr() } : t))
    );
    registrarHistorial(id, `${etiqueta} actualizado a "${valor}".`);
  }

  function guardarGestion() {
    if (!seleccionado) return;
    setTickets((prev) =>
      prev.map((t) =>
        t.id === seleccionado.id
          ? { ...t, respuestaEstudiante: respuestaLocal, novedadArea: novedadLocal, actualizado: nowStr() }
          : t
      )
    );
    registrarHistorial(seleccionado.id, "Se actualizó la respuesta al estudiante y/o la novedad de área.");
    setGuardadoOk(true);
    setTimeout(() => setGuardadoOk(false), 2000);
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
      <div className="flex items-center justify-between px-6 py-3" style={{ background: "#12192A", borderBottom: "1px solid #232C40" }}>
        <div className="flex items-center gap-3">
          <div style={{ width: 30, height: 30, borderRadius: 8, background: "#2F6FED", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Inbox size={16} color="#FFFFFF" />
          </div>
          <div>
            <div style={{ color: "#F7F8FA", fontWeight: 700, fontSize: 14, lineHeight: 1.1 }}>Mesa de Ayuda</div>
            <div style={{ color: "#7A8698", fontSize: 11 }}>Panel de gestión de tickets</div>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-lg" style={{ background: "#1C2536" }}>
            <div style={{ width: 22, height: 22, borderRadius: 999, background: "#2F6FED", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: "#fff", fontWeight: 700 }}>
              {agente.nombre.split(" ").map((p) => p[0]).slice(0, 2).join("")}
            </div>
            <span style={{ color: "#E5E9F0", fontSize: 13 }}>{agente.nombre}</span>
          </div>
          <button onClick={onLogout} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold"
            style={{ color: "#F7F8FA", background: "#2A3542" }}>
            <LogOut size={13} /> Salir
          </button>
        </div>
      </div>

      <div className="px-4 sm:px-6 py-5" style={{ maxWidth: 1600, margin: "0 auto" }}>
        <div className="flex gap-2 mb-5">
          <button onClick={() => setTab("tickets")}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{ background: tab === "tickets" ? "#12192A" : "#FFFFFF", color: tab === "tickets" ? "#FFFFFF" : "#5A6577", border: "1px solid #E2E6EC" }}>
            Bandeja de tickets
          </button>
          <button onClick={() => setTab("reportes")}
            className="px-4 py-2 rounded-lg text-sm font-semibold"
            style={{ background: tab === "reportes" ? "#12192A" : "#FFFFFF", color: tab === "reportes" ? "#FFFFFF" : "#5A6577", border: "1px solid #E2E6EC" }}>
            Reportes
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
                <div className="grid grid-cols-3 gap-1.5">
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
                            <button onClick={() => setMostrarPassword((v) => !v)} style={{ color: "#2F6FED" }}>
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
                    <div className="text-sm p-3 rounded-lg" style={{ background: "#F7F8FA", color: "#1B2430" }}>
                      {seleccionado.descripcion}
                    </div>

                    <div className="mt-3">
                      <div className="flex items-center gap-1.5 text-[11px] mb-2" style={{ color: "#9AA4B2" }}>
                        <Paperclip size={12} /> Evidencia adjunta
                      </div>
                      {seleccionado.evidencia ? (
                        <img src={seleccionado.evidencia.dataUrl} alt="Evidencia" className="rounded-lg" style={{ maxHeight: 160 }} />
                      ) : (
                        <div className="text-xs" style={{ color: "#B7BFCB" }}>El estudiante no adjuntó evidencia.</div>
                      )}
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
                        <div key={i} className="text-xs p-2.5 rounded-lg" style={{ background: "#F7F8FA" }}>
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
                  <div className="px-5 py-4" style={{ background: "#12192A" }}>
                    <div className="text-xs font-bold" style={{ color: "#F7F8FA", letterSpacing: "0.05em" }}>GESTIÓN INTERNA</div>
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
                          {ASESORES.map((a) => <option key={a}>{a}</option>)}
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="p-5 space-y-4">
                    <div>
                      <label className="flex items-center gap-1.5 text-[11px] font-semibold mb-1" style={{ color: "#5A6577" }}>
                        <MessageSquare size={12} /> RESPUESTA AL ESTUDIANTE
                      </label>
                      <div className="text-[11px] mb-1.5" style={{ color: "#9AA4B2" }}>Único mensaje visible para el estudiante.</div>
                      <textarea rows={3} value={respuestaLocal} onChange={(e) => setRespuestaLocal(e.target.value)}
                        placeholder="Ej: Tu caso está siendo gestionado, en breve tendrás una solución."
                        className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                    </div>

                    <div>
                      <label className="text-[11px] font-semibold mb-1 block" style={{ color: "#5A6577" }}>NOVEDAD ÁREA</label>
                      <div className="text-[11px] mb-1.5" style={{ color: "#9AA4B2" }}>Nota interna para remitir el caso entre áreas. No visible para el estudiante.</div>
                      <textarea rows={3} value={novedadLocal} onChange={(e) => setNovedadLocal(e.target.value)}
                        placeholder="Ej: Se remite a Prizma para validar el bloqueo de la cuenta."
                        className="w-full px-3 py-2 rounded-lg text-sm" style={{ border: "1px solid #DDE2E9" }} />
                    </div>

                    <button onClick={guardarGestion}
                      className="w-full flex items-center justify-center gap-2 py-2.5 rounded-lg text-sm font-semibold"
                      style={{ background: guardadoOk ? "#2FAE6B" : "#2F6FED", color: "#FFFFFF" }}>
                      <Save size={14} /> {guardadoOk ? "Guardado" : "Guardar cambios"}
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "reportes" && <ModuloReportes tickets={tickets} />}
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

// ---------- Módulo: Dashboard Ejecutivo de Reportes ----------
function ModuloReportes({ tickets }) {
  const categorias = Object.keys(CATEGORIAS_ESTUDIANTE);
  const asesoresReales = ASESORES.filter((a) => a !== "Sin asignar");

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
    const rows = [["Ticket", "Estudiante", "Identificacion", "Categoria", "Subcategoria", "Area", "Asesor", "Prioridad", "Estado", "Creado"]];
    ticketsFiltrados.forEach((t) => rows.push([
      "TCK-" + String(t.id).padStart(4, "0"), t.nombre, t.codigo, t.categoria, t.subcategoria || "",
      t.area, t.asesor, t.prioridad, t.estado, t.creado,
    ]));
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
          <Users size={15} color="#2F6FED" /> Gestión por asesor
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
                <Bar dataKey="asignados" name="Asignados" fill="#2F6FED" radius={[0, 4, 4, 0]} />
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
          <Building2 size={15} color="#2F6FED" /> Gestión por área
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
            <Activity size={15} color="#2F6FED" /> Tickets creados por día
          </div>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <LineChart data={ticketsDiarios}>
                <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" />
                <XAxis dataKey="fecha" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                <Tooltip />
                <Line type="monotone" dataKey="cantidad" name="Creados" stroke="#2F6FED" strokeWidth={2} dot={{ r: 3 }} />
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
            <PieIcon size={15} color="#2F6FED" /> Distribución por categoría
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
            <BarChart3 size={15} color="#2F6FED" /> Distribución por prioridad
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
            const color = t.tipo === "up" ? "#2FAE6B" : t.tipo === "down" ? "#D64545" : "#2F6FED";
            return (
              <div key={i} className="flex items-start gap-2 text-sm p-2.5 rounded-lg" style={{ background: "#F7F8FA" }}>
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

// ---------- App principal ----------
export default function App() {
  const [vista, setVista] = useState("home"); // 'home' | 'estudiante' | 'soporte-login' | 'soporte-dashboard'
  const [agente, setAgente] = useState(null);
  const [tickets, setTickets] = useState(SEED_TICKETS);

  return (
    <div style={{ fontFamily: "Inter, ui-sans-serif, system-ui, sans-serif" }}>
      {vista === "home" && <Portada onSelect={setVista} />}
      {vista === "estudiante" && (
        <VistaEstudiante tickets={tickets} setTickets={setTickets} onBack={() => setVista("home")} />
      )}
      {vista === "soporte" && (
        <LoginSoporte onLogin={(a) => { setAgente(a); setVista("soporte-dashboard"); }} onBack={() => setVista("home")} />
      )}
      {vista === "soporte-dashboard" && agente && (
        <DashboardSoporte
          agente={agente}
          tickets={tickets}
          setTickets={setTickets}
          onLogout={() => { setAgente(null); setVista("home"); }}
        />
      )}
    </div>
  );
}
