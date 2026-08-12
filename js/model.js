// ============================================================
// MODELO DE DATOS
// Cuatro entidades: alumnos, cursos, ediciones, inscripciones
// + promociones (reglas curso disparador → edición de regalo)
// ============================================================

export function baseVacia() {
  return {
    schemaVersion: 1,
    meta: {
      lastSavedBy: "",
      lastSavedAt: null,
      rev: 0,
      passwordHash: null
    },
    alumnos: [],
    cursos: [],
    ediciones: [],
    inscripciones: [],
    promociones: []
  };
}

export function validarBase(obj) {
  if (!obj || typeof obj !== "object") return false;
  return ["alumnos", "cursos", "ediciones", "inscripciones"].every(k => Array.isArray(obj[k]));
}

// Migra/completa campos que pudieran faltar en archivos viejos
export function normalizarBase(db) {
  const base = baseVacia();
  db.meta = Object.assign(base.meta, db.meta || {});
  db.promociones = db.promociones || [];
  db.schemaVersion = db.schemaVersion || 1;
  return db;
}

const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random().toString(36).slice(2));

export const hoyISO = () => new Date().toISOString().slice(0, 10);

// ---------- utilidades de texto ----------
export function normalizar(s) {
  return (s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

export function nombreCompleto(a) {
  return [a.nombres, a.apellidoPaterno, a.apellidoMaterno].filter(Boolean).join(" ").trim();
}

export function nombreLista(a) {
  // "Apellido Paterno Apellido Materno, Nombres" para listas alfabéticas
  const apellidos = [a.apellidoPaterno, a.apellidoMaterno].filter(Boolean).join(" ");
  return apellidos ? `${apellidos}, ${a.nombres || ""}`.trim() : (a.nombres || "(sin nombre)");
}

// ---------- búsqueda ----------
export function buscarAlumnos(db, q) {
  const nq = normalizar(q);
  if (!nq) return [];
  const partes = nq.split(" ");
  return db.alumnos.filter(a => {
    const texto = normalizar(nombreCompleto(a) + " " + (a.email || "") + " " + (a.telefono || ""));
    return partes.every(p => texto.includes(p));
  });
}

export function buscarCursos(db, q) {
  const nq = normalizar(q);
  if (!nq) return [];
  return db.cursos.filter(c => normalizar(c.nombre).includes(nq));
}

// ---------- duplicados ----------
export function detectarDuplicados(db, { email, nombres, apellidoPaterno, apellidoMaterno }, excluirId = null) {
  const resultados = [];
  const ne = normalizar(email);
  const nnombre = normalizar([nombres, apellidoPaterno, apellidoMaterno].filter(Boolean).join(" "));
  for (const a of db.alumnos) {
    if (a.id === excluirId) continue;
    const mismoEmail = ne && normalizar(a.email) === ne;
    const mismoNombre = nnombre && normalizar(nombreCompleto(a)) === nnombre;
    if (mismoEmail || mismoNombre) {
      resultados.push({ alumno: a, porEmail: mismoEmail, porNombre: mismoNombre });
    }
  }
  return resultados;
}

// ---------- alumnos ----------
export function agregarAlumno(db, datos) {
  const a = {
    id: uid(),
    nombres: (datos.nombres || "").trim(),
    apellidoPaterno: (datos.apellidoPaterno || "").trim(),
    apellidoMaterno: (datos.apellidoMaterno || "").trim(),
    email: (datos.email || "").trim(),
    telefono: (datos.telefono || "").trim(),
    nota: (datos.nota || "").trim(),
    revisar: !!datos.revisar,
    createdAt: new Date().toISOString()
  };
  db.alumnos.push(a);
  return a;
}

export function actualizarAlumno(db, id, datos) {
  const a = db.alumnos.find(x => x.id === id);
  if (a) Object.assign(a, datos);
  return a;
}

export function eliminarAlumno(db, id) {
  db.inscripciones = db.inscripciones.filter(i => i.alumnoId !== id);
  db.alumnos = db.alumnos.filter(a => a.id !== id);
}

// ---------- cursos ----------
export function agregarCurso(db, { nombre, tipo, descripcion }) {
  const c = {
    id: uid(),
    nombre: (nombre || "").trim(),
    tipo: tipo === "video" ? "video" : "vivo",
    descripcion: (descripcion || "").trim()
  };
  db.cursos.push(c);
  return c;
}

export function cursoDuplicado(db, nombre, excluirId = null) {
  const nn = normalizar(nombre);
  return db.cursos.find(c => c.id !== excluirId && normalizar(c.nombre) === nn) || null;
}

export function eliminarCurso(db, id) {
  const edIds = db.ediciones.filter(e => e.cursoId === id).map(e => e.id);
  db.inscripciones = db.inscripciones.filter(i => !edIds.includes(i.edicionId));
  db.ediciones = db.ediciones.filter(e => e.cursoId !== id);
  db.promociones = db.promociones.filter(p => p.triggerCursoId !== id && !edIds.includes(p.giftEdicionId));
  db.cursos = db.cursos.filter(c => c.id !== id);
}

// ---------- ediciones ----------
export function agregarEdicion(db, { cursoId, periodo, nota }) {
  const e = {
    id: uid(),
    cursoId,
    periodo: (periodo || "").trim(),
    nota: (nota || "").trim()
  };
  db.ediciones.push(e);
  return e;
}

export function eliminarEdicion(db, id) {
  db.inscripciones = db.inscripciones.filter(i => i.edicionId !== id);
  db.promociones = db.promociones.filter(p => p.giftEdicionId !== id);
  db.ediciones = db.ediciones.filter(e => e.id !== id);
}

export function edicionesDeCurso(db, cursoId) {
  return db.ediciones.filter(e => e.cursoId === cursoId);
}

export function etiquetaEdicion(db, edicionId) {
  const e = db.ediciones.find(x => x.id === edicionId);
  if (!e) return "(edición eliminada)";
  const c = db.cursos.find(x => x.id === e.cursoId);
  return `${c ? c.nombre : "?"} · ${e.periodo || "sin fecha"}`;
}

// ---------- inscripciones + promociones ----------
export function yaInscrito(db, alumnoId, edicionId) {
  return db.inscripciones.some(i => i.alumnoId === alumnoId && i.edicionId === edicionId);
}

// Inscribe y aplica promociones. Devuelve { inscripcion, promos: [inscripcion...] }
export function inscribir(db, { alumnoId, edicionId, fecha, nota, esPromo = false, promoDe = null }) {
  const i = {
    id: uid(),
    alumnoId,
    edicionId,
    fecha: fecha || hoyISO(),
    nota: (nota || "").trim(),
    promo: esPromo,
    promoDe
  };
  db.inscripciones.push(i);

  const promos = [];
  if (!esPromo) {
    const edicion = db.ediciones.find(e => e.id === edicionId);
    if (edicion) {
      const reglas = db.promociones.filter(p => p.activa && p.triggerCursoId === edicion.cursoId);
      for (const regla of reglas) {
        if (!yaInscrito(db, alumnoId, regla.giftEdicionId)) {
          const { inscripcion } = inscribir(db, {
            alumnoId,
            edicionId: regla.giftEdicionId,
            fecha: i.fecha,
            nota: "Promoción automática",
            esPromo: true,
            promoDe: i.id
          });
          promos.push(inscripcion);
        }
      }
    }
  }
  return { inscripcion: i, promos };
}

export function eliminarInscripcion(db, id) {
  // elimina también las inscripciones-regalo que nacieron de ésta
  db.inscripciones = db.inscripciones.filter(i => i.id !== id && i.promoDe !== id);
}

export function inscripcionesDeAlumno(db, alumnoId) {
  return db.inscripciones.filter(i => i.alumnoId === alumnoId);
}

export function inscritosDeEdicion(db, edicionId) {
  return db.inscripciones.filter(i => i.edicionId === edicionId);
}

// ---------- promociones ----------
export function agregarPromocion(db, { triggerCursoId, giftEdicionId }) {
  const p = { id: uid(), triggerCursoId, giftEdicionId, activa: true };
  db.promociones.push(p);
  return p;
}

export function eliminarPromocion(db, id) {
  db.promociones = db.promociones.filter(p => p.id !== id);
}

// ---------- contraseña ----------
export async function hashContrasena(texto) {
  const data = new TextEncoder().encode("alumnos-priosan:" + texto);
  const buf = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}
