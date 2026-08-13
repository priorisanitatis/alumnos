// ============================================================
// EXPORTACIÓN A HOJAS DE CÁLCULO
//
// Genera CSV con BOM UTF-8, que Google Sheets, Excel y Numbers
// abren directamente. Aquí solo se CONSTRUYE el contenido; quien
// decide dónde va a parar (Google Drive o Descargas) es app.js.
// ============================================================

import { nombreCompleto } from "./model.js";

function celda(v) {
  const s = String(v ?? "");
  return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function armarCSV(filas) {
  return "\ufeff" + filas.map(f => f.map(celda).join(",")).join("\r\n");
}

const limpiaNombre = s => s.replace(/[^\wáéíóúñÁÉÍÓÚÑ -]/gi, "").trim().replace(/\s+/g, "_");

// ---------- constructores: devuelven { nombre, contenido } ----------

export function construirAlumnos(db) {
  const filas = [["Apellido paterno", "Apellido materno", "Nombre(s)", "Email", "Teléfono", "Nota", "Cursos tomados"]];
  const alumnos = [...db.alumnos].sort((a, b) =>
    (a.apellidoPaterno + a.apellidoMaterno + a.nombres).localeCompare(b.apellidoPaterno + b.apellidoMaterno + b.nombres, "es"));
  for (const a of alumnos) {
    const n = db.inscripciones.filter(i => i.alumnoId === a.id).length;
    filas.push([a.apellidoPaterno, a.apellidoMaterno, a.nombres, a.email, a.telefono, a.nota, n]);
  }
  return { nombre: "alumnos_todos.csv", contenido: armarCSV(filas) };
}

export function construirInscripciones(db) {
  const filas = [["Alumno", "Email", "Teléfono", "Curso", "Edición", "Fecha de inscripción", "Promoción", "Nota"]];
  for (const i of db.inscripciones) {
    const a = db.alumnos.find(x => x.id === i.alumnoId);
    const e = db.ediciones.find(x => x.id === i.edicionId);
    const c = e ? db.cursos.find(x => x.id === e.cursoId) : null;
    filas.push([
      a ? nombreCompleto(a) : "(eliminado)",
      a?.email || "", a?.telefono || "",
      c?.nombre || "?", e?.periodo || "?",
      i.fecha, i.promo ? "Sí" : "", i.nota
    ]);
  }
  return { nombre: "inscripciones_todas.csv", contenido: armarCSV(filas) };
}

export function construirEdicion(db, edicionId) {
  const e = db.ediciones.find(x => x.id === edicionId);
  const c = e ? db.cursos.find(x => x.id === e.cursoId) : null;
  const filas = [["Apellido paterno", "Apellido materno", "Nombre(s)", "Email", "Teléfono", "Fecha de inscripción", "Promoción", "Nota"]];
  for (const i of db.inscripciones.filter(x => x.edicionId === edicionId)) {
    const a = db.alumnos.find(x => x.id === i.alumnoId);
    filas.push([a?.apellidoPaterno || "", a?.apellidoMaterno || "", a?.nombres || "(eliminado)",
      a?.email || "", a?.telefono || "", i.fecha, i.promo ? "Sí" : "", i.nota]);
  }
  return {
    nombre: limpiaNombre((c?.nombre || "curso") + "_" + (e?.periodo || "")) + ".csv",
    contenido: armarCSV(filas)
  };
}

export function construirCurso(db, cursoId) {
  const c = db.cursos.find(x => x.id === cursoId);
  const filas = [["Edición", "Apellido paterno", "Apellido materno", "Nombre(s)", "Email", "Teléfono", "Fecha de inscripción", "Promoción", "Nota"]];
  for (const e of db.ediciones.filter(x => x.cursoId === cursoId)) {
    for (const i of db.inscripciones.filter(x => x.edicionId === e.id)) {
      const a = db.alumnos.find(x => x.id === i.alumnoId);
      filas.push([e.periodo, a?.apellidoPaterno || "", a?.apellidoMaterno || "", a?.nombres || "(eliminado)",
        a?.email || "", a?.telefono || "", i.fecha, i.promo ? "Sí" : "", i.nota]);
    }
  }
  return {
    nombre: limpiaNombre(c?.nombre || "curso") + "_todas_ediciones.csv",
    contenido: armarCSV(filas)
  };
}

// ---------- constancias ----------
// Lista pensada para generar constancias en lote (Canva "Crear en lote") y
// para la app de envíos: nombre completo en UNA sola celda, orden alfabético
// por nombre(s), y el curso y la fecha repetidos en cada fila para que la
// plantilla pueda tomarlos como campos.

// Arregla la ortografía del nombre SOLO si viene mal capturado (todo en
// minúsculas o TODO EN MAYÚSCULAS). Si ya trae mayúsculas y minúsculas
// mezcladas, se respeta tal cual: quien lo escribió sabía lo que hacía
// (por ejemplo "de la Cruz" o "McKenna").
const PARTICULAS = new Set(["de", "del", "la", "las", "los", "y", "e", "da", "van", "von", "di"]);

export function nombreParaConstancia(s) {
  const t = (s || "").trim();
  if (!t) return "";
  const todoMinusculas = t === t.toLowerCase();
  const todoMayusculas = t === t.toUpperCase();
  if (!todoMinusculas && !todoMayusculas) return t;
  return t.toLowerCase().split(/\s+/)
    .map((p, i) => (i > 0 && PARTICULAS.has(p)) ? p : p.charAt(0).toUpperCase() + p.slice(1))
    .join(" ");
}

function inscritosOrdenadosPorNombre(db, edicionId) {
  return db.inscripciones
    .filter(i => i.edicionId === edicionId)
    .map(i => ({ i, a: db.alumnos.find(x => x.id === i.alumnoId) }))
    .filter(x => x.a)
    .sort((x, y) => nombreCompleto(x.a).localeCompare(nombreCompleto(y.a), "es", { sensitivity: "base" }));
}

export function construirConstancias(db, edicionId) {
  const e = db.ediciones.find(x => x.id === edicionId);
  const c = e ? db.cursos.find(x => x.id === e.cursoId) : null;
  const filas = [["#", "Nombre completo", "Email", "Teléfono", "Curso", "Fecha"]];
  inscritosOrdenadosPorNombre(db, edicionId).forEach(({ a }, idx) => {
    filas.push([idx + 1, nombreParaConstancia(nombreCompleto(a)), a.email || "", a.telefono || "",
      c?.nombre || "", e?.periodo || ""]);
  });
  return {
    nombre: "constancias_" + limpiaNombre((c?.nombre || "curso") + "_" + (e?.periodo || "")) + ".csv",
    contenido: armarCSV(filas)
  };
}

// Solo los nombres, uno por línea, para copiar y pegar directamente.
export function nombresParaPegar(db, edicionId) {
  return inscritosOrdenadosPorNombre(db, edicionId)
    .map(({ a }) => nombreParaConstancia(nombreCompleto(a)))
    .join("\n");
}

export function construirRespaldoJSON(db) {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return {
    nombre: `respaldo-alumnos-${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.json`,
    contenido: JSON.stringify(db, null, 2)
  };
}

// ---------- entrega por descarga (carpeta Descargas) ----------

export function descargar({ nombre, contenido }) {
  const tipo = nombre.endsWith(".json") ? "application/json" : "text/csv;charset=utf-8";
  const blob = new Blob([contenido], { type: tipo });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  a.click();
  URL.revokeObjectURL(a.href);
}
