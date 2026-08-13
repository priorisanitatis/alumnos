// ============================================================
// IMPORTACIÓN DE ALUMNOS DESDE CSV (respuestas de Google Forms)
//
// Flujo: leer archivo → detectar columnas → ANALIZAR (sin tocar nada)
// → el usuario revisa la vista previa → APLICAR.
// Nunca se modifica la base hasta que el usuario confirma.
// ============================================================

import { normalizar, nombreCompleto, agregarAlumno, inscribir, yaInscrito } from "./model.js";

// ---------- lectura de CSV ----------

// Detecta si el archivo usa coma, punto y coma o tabulador.
function detectarDelimitador(texto) {
  const primeraLinea = texto.split(/\r?\n/)[0] || "";
  let fuera = true, cuenta = { ",": 0, ";": 0, "\t": 0 };
  for (const ch of primeraLinea) {
    if (ch === '"') fuera = !fuera;
    else if (fuera && cuenta[ch] !== undefined) cuenta[ch]++;
  }
  return Object.keys(cuenta).reduce((a, b) => (cuenta[b] > cuenta[a] ? b : a), ",");
}

export function parsearCSV(texto) {
  if (texto.charCodeAt(0) === 0xFEFF) texto = texto.slice(1); // quita el BOM
  const D = detectarDelimitador(texto);
  const filas = [];
  let fila = [], campo = "", enComillas = false;

  for (let i = 0; i < texto.length; i++) {
    const ch = texto[i];
    if (enComillas) {
      if (ch === '"') {
        if (texto[i + 1] === '"') { campo += '"'; i++; }
        else enComillas = false;
      } else campo += ch;
    } else if (ch === '"') {
      enComillas = true;
    } else if (ch === D) {
      fila.push(campo); campo = "";
    } else if (ch === "\n") {
      fila.push(campo); campo = "";
      if (fila.some(c => c.trim() !== "")) filas.push(fila);
      fila = [];
    } else if (ch !== "\r") {
      campo += ch;
    }
  }
  fila.push(campo);
  if (fila.some(c => c.trim() !== "")) filas.push(fila);

  return filas.map(f => f.map(c => c.trim()));
}

// ---------- detección automática de columnas ----------

const REGLAS = [
  ["apellidoPaterno", /paterno/],
  ["apellidoMaterno", /materno/],
  ["email", /correo|e-?mail/],
  ["telefono", /tel|cel|whats|movil|m[oó]vil/],
  ["nombreCompleto", /nombre\s*completo|nombre\s*y\s*apellidos/],
  ["nombres", /nombre/]
];

// Devuelve { campo: índiceDeColumna } a partir de los encabezados.
export function detectarColumnas(encabezados) {
  const mapeo = {};
  encabezados.forEach((h, i) => {
    const n = normalizar(h);
    for (const [campo, patron] of REGLAS) {
      if (mapeo[campo] !== undefined) continue;
      if (patron.test(n)) {
        // "nombre completo" gana sobre "nombres" para la misma columna
        if (campo === "nombres" && /completo/.test(n)) continue;
        mapeo[campo] = i;
        break;
      }
    }
  });
  // Si detectó nombre completo Y nombres por separado, manda el desglosado.
  if (mapeo.nombres !== undefined && mapeo.apellidoPaterno !== undefined) {
    delete mapeo.nombreCompleto;
  }
  return mapeo;
}

// ---------- separación de nombre completo (solo si hace falta) ----------
// Convención mexicana: los dos últimos bloques son los apellidos.
// Es imperfecto a propósito: las filas así se marcan para revisión.
const PARTICULAS = new Set(["de", "del", "la", "las", "los", "y", "san", "santa"]);

export function separarNombre(completo) {
  const palabras = (completo || "").trim().split(/\s+/).filter(Boolean);
  if (palabras.length === 0) return { nombres: "", apellidoPaterno: "", apellidoMaterno: "", dudoso: true };
  if (palabras.length === 1) return { nombres: palabras[0], apellidoPaterno: "", apellidoMaterno: "", dudoso: true };
  if (palabras.length === 2) return { nombres: palabras[0], apellidoPaterno: palabras[1], apellidoMaterno: "", dudoso: false };

  // Agrupa partículas con la palabra que les sigue: "de la Cruz" = un apellido
  const bloques = [];
  for (let i = 0; i < palabras.length; i++) {
    if (PARTICULAS.has(normalizar(palabras[i])) && i < palabras.length - 1) {
      let bloque = palabras[i];
      while (i + 1 < palabras.length && PARTICULAS.has(normalizar(palabras[i + 1]))) { bloque += " " + palabras[++i]; }
      bloque += " " + palabras[++i];
      bloques.push(bloque);
    } else {
      bloques.push(palabras[i]);
    }
  }
  if (bloques.length < 3) {
    return { nombres: bloques[0] || "", apellidoPaterno: bloques[1] || "", apellidoMaterno: "", dudoso: true };
  }
  const apellidoMaterno = bloques.pop();
  const apellidoPaterno = bloques.pop();
  return {
    nombres: bloques.join(" "),
    apellidoPaterno,
    apellidoMaterno,
    dudoso: bloques.length > 2 // 3+ nombres de pila: mejor que lo revise
  };
}

// ---------- validación ----------
const RE_EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
export const emailValido = e => RE_EMAIL.test((e || "").trim());

// ---------- análisis (NO modifica la base) ----------
//
// Cada fila se clasifica en:
//   "nuevo"      → alumno que no existe: se creará
//   "actualiza"  → ya existe y el archivo aporta datos que le faltaban
//   "conflicto"  → ya existe pero el archivo trae un dato DISTINTO
//   "sin-cambio" → ya existe y no aporta nada nuevo
//   "error"      → no se puede usar (sin nombre, email inválido…)
//   "repetido"   → aparece dos veces dentro del mismo archivo

export function analizar(db, filas, mapeo) {
  const encabezado = filas[0];
  const cuerpo = filas.slice(1);
  const resultado = [];
  const vistosEmail = new Map();
  const vistosNombre = new Map();

  const val = (fila, campo) => (mapeo[campo] !== undefined ? (fila[mapeo[campo]] || "").trim() : "");

  cuerpo.forEach((fila, idx) => {
    let nombres = val(fila, "nombres");
    let apellidoPaterno = val(fila, "apellidoPaterno");
    let apellidoMaterno = val(fila, "apellidoMaterno");
    let dudoso = false;

    // Si solo vino el nombre completo, hay que separarlo (imperfecto)
    if (!nombres && mapeo.nombreCompleto !== undefined) {
      const p = separarNombre(val(fila, "nombreCompleto"));
      nombres = p.nombres; apellidoPaterno = p.apellidoPaterno; apellidoMaterno = p.apellidoMaterno;
      dudoso = p.dudoso;
    }

    let email = val(fila, "email").toLowerCase();
    const telefono = val(fila, "telefono");
    const fil = { linea: idx + 2, estado: "nuevo", motivo: "", advertencia: "", alumnoId: null, cambios: [] };

    // --- único error de verdad: sin nombre no hay a quién registrar ---
    if (!nombres && !apellidoPaterno) {
      fil.entrada = { nombres, apellidoPaterno, apellidoMaterno, email, telefono, revisar: true };
      fil.estado = "error"; fil.motivo = "Sin nombre";
      resultado.push(fil); return;
    }

    // Un email mal escrito NO descarta a la persona: se importa sin email
    // y queda marcada para revisión, para no perder a nadie por un dedazo.
    if (email && !emailValido(email)) {
      fil.advertencia = `Email mal escrito («${email}»): se importa SIN email y marcado para revisar`;
      email = "";
      dudoso = true;
    }

    const entrada = { nombres, apellidoPaterno, apellidoMaterno, email, telefono, revisar: dudoso };
    fil.entrada = entrada;

    // --- repetidos dentro del propio archivo ---
    const claveNombre = normalizar([nombres, apellidoPaterno, apellidoMaterno].join(" "));
    if (email && vistosEmail.has(email)) {
      fil.estado = "repetido"; fil.motivo = "Mismo email que la línea " + vistosEmail.get(email);
      resultado.push(fil); return;
    }
    // Mismo nombre repetido en el archivo: casi siempre alguien que llenó el
    // formulario dos veces. Se marca para que tú decidas, en vez de duplicarlo.
    if (claveNombre && vistosNombre.has(claveNombre)) {
      fil.estado = "repetido";
      fil.motivo = "Mismo nombre que la línea " + vistosNombre.get(claveNombre) +
        " (si de verdad son dos personas distintas, captura la segunda a mano)";
      resultado.push(fil); return;
    }
    if (email) vistosEmail.set(email, fil.linea);
    if (claveNombre) vistosNombre.set(claveNombre, fil.linea);

    // --- ¿ya existe en la base? primero por email, luego por nombre ---
    let existente = email ? db.alumnos.find(a => normalizar(a.email) === normalizar(email)) : null;
    let porNombre = false;
    if (!existente && claveNombre) {
      existente = db.alumnos.find(a => normalizar(nombreCompleto(a)) === claveNombre);
      porNombre = !!existente;
    }

    if (!existente) { resultado.push(fil); return; }

    fil.alumnoId = existente.id;
    fil.encontradoPor = porNombre ? "nombre" : "email";

    // ¿qué aporta el archivo?
    for (const campo of ["email", "telefono", "apellidoMaterno"]) {
      const nuevo = entrada[campo];
      const viejo = (existente[campo] || "").trim();
      if (!nuevo) continue;
      if (!viejo) fil.cambios.push({ campo, de: "", a: nuevo, tipo: "rellena" });
      else if (normalizar(viejo) !== normalizar(nuevo)) fil.cambios.push({ campo, de: viejo, a: nuevo, tipo: "distinto" });
    }

    if (!fil.cambios.length) { fil.estado = "sin-cambio"; fil.motivo = "Ya está igual en la base"; }
    else if (fil.cambios.some(c => c.tipo === "distinto")) { fil.estado = "conflicto"; }
    else { fil.estado = "actualiza"; }

    resultado.push(fil);
  });

  return { encabezado, filas: resultado, resumen: contar(resultado) };
}

function contar(filas) {
  const r = { nuevo: 0, actualiza: 0, conflicto: 0, "sin-cambio": 0, error: 0, repetido: 0 };
  filas.forEach(f => r[f.estado]++);
  return r;
}

// ---------- aplicación (SÍ modifica la base) ----------
//
// opciones:
//   preferirNuevos: en los conflictos, sobreescribir el dato viejo
//   edicionId: si viene, inscribe a todos los importados en esa edición

export function aplicar(db, analisis, opciones = {}) {
  const { preferirNuevos = false, edicionId = null } = opciones;
  const hecho = { creados: 0, actualizados: 0, inscritos: 0, promos: 0, omitidos: 0 };
  const afectados = [];

  for (const fil of analisis.filas) {
    if (fil.estado === "error" || fil.estado === "repetido") { hecho.omitidos++; continue; }

    let alumno;
    if (fil.estado === "nuevo") {
      alumno = agregarAlumno(db, fil.entrada);
      hecho.creados++;
    } else {
      alumno = db.alumnos.find(a => a.id === fil.alumnoId);
      if (!alumno) { hecho.omitidos++; continue; }
      let tocado = false;
      for (const c of fil.cambios) {
        if (c.tipo === "rellena" || preferirNuevos) { alumno[c.campo] = c.a; tocado = true; }
      }
      if (tocado) hecho.actualizados++;
    }
    afectados.push(alumno);

    if (edicionId && !yaInscrito(db, alumno.id, edicionId)) {
      const { promos } = inscribir(db, { alumnoId: alumno.id, edicionId, nota: "Alta por formulario" });
      hecho.inscritos++;
      hecho.promos += promos.length;
    }
  }
  return { ...hecho, afectados };
}
