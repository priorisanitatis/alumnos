// ============================================================
// EXPORTACIÓN A EXCEL (CSV con BOM UTF-8, se abre directo en Excel)
// ============================================================

import { nombreCompleto, etiquetaEdicion } from "./model.js";

function celda(v) {
  const s = String(v ?? "");
  return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function descargarCSV(nombreArchivo, filas) {
  const contenido = "\ufeff" + filas.map(f => f.map(celda).join(",")).join("\r\n");
  const blob = new Blob([contenido], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = nombreArchivo;
  a.click();
  URL.revokeObjectURL(a.href);
}

const limpiaNombre = s => s.replace(/[^\wáéíóúñÁÉÍÓÚÑ -]/gi, "").trim().replace(/\s+/g, "_");

export function exportarAlumnos(db) {
  const filas = [["Apellido paterno", "Apellido materno", "Nombre(s)", "Email", "Teléfono", "Nota", "Cursos tomados"]];
  const alumnos = [...db.alumnos].sort((a, b) =>
    (a.apellidoPaterno + a.apellidoMaterno + a.nombres).localeCompare(b.apellidoPaterno + b.apellidoMaterno + b.nombres, "es"));
  for (const a of alumnos) {
    const n = db.inscripciones.filter(i => i.alumnoId === a.id).length;
    filas.push([a.apellidoPaterno, a.apellidoMaterno, a.nombres, a.email, a.telefono, a.nota, n]);
  }
  descargarCSV("alumnos_todos.csv", filas);
}

export function exportarInscripciones(db) {
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
  descargarCSV("inscripciones_todas.csv", filas);
}

export function exportarEdicion(db, edicionId) {
  const e = db.ediciones.find(x => x.id === edicionId);
  const c = e ? db.cursos.find(x => x.id === e.cursoId) : null;
  const filas = [["Apellido paterno", "Apellido materno", "Nombre(s)", "Email", "Teléfono", "Fecha de inscripción", "Promoción", "Nota"]];
  for (const i of db.inscripciones.filter(x => x.edicionId === edicionId)) {
    const a = db.alumnos.find(x => x.id === i.alumnoId);
    filas.push([a?.apellidoPaterno || "", a?.apellidoMaterno || "", a?.nombres || "(eliminado)",
      a?.email || "", a?.telefono || "", i.fecha, i.promo ? "Sí" : "", i.nota]);
  }
  descargarCSV(limpiaNombre((c?.nombre || "curso") + "_" + (e?.periodo || "")) + ".csv", filas);
}

export function exportarCurso(db, cursoId) {
  const c = db.cursos.find(x => x.id === cursoId);
  const filas = [["Edición", "Apellido paterno", "Apellido materno", "Nombre(s)", "Email", "Teléfono", "Fecha de inscripción", "Promoción", "Nota"]];
  for (const e of db.ediciones.filter(x => x.cursoId === cursoId)) {
    for (const i of db.inscripciones.filter(x => x.edicionId === e.id)) {
      const a = db.alumnos.find(x => x.id === i.alumnoId);
      filas.push([e.periodo, a?.apellidoPaterno || "", a?.apellidoMaterno || "", a?.nombres || "(eliminado)",
        a?.email || "", a?.telefono || "", i.fecha, i.promo ? "Sí" : "", i.nota]);
    }
  }
  descargarCSV(limpiaNombre(c?.nombre || "curso") + "_todas_ediciones.csv", filas);
}

export function exportarRespaldoJSON(db) {
  const blob = new Blob([JSON.stringify(db, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  const d = new Date(), p = n => String(n).padStart(2, "0");
  a.href = URL.createObjectURL(blob);
  a.download = `respaldo-alumnos-${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}
