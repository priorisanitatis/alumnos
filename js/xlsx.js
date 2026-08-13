// ============================================================
// LECTOR DE ARCHIVOS EXCEL (.xlsx)
//
// Un .xlsx es en realidad un ZIP con archivos XML adentro. Aquí se
// abre ese ZIP y se leen las dos piezas que importan:
//   xl/sharedStrings.xml         → el texto de las celdas
//   xl/worksheets/sheet1.xml     → la primera hoja
//
// No usa ninguna librería externa: descomprime con DecompressionStream,
// que ya viene en Chrome. Devuelve lo mismo que el lector de CSV:
// un arreglo de filas, cada una un arreglo de textos.
// ============================================================

// ---------- ZIP ----------

function leerU16(v, p) { return v.getUint16(p, true); }
function leerU32(v, p) { return v.getUint32(p, true); }

// Localiza el índice del ZIP (End Of Central Directory) buscándolo desde el final.
function buscarEOCD(vista) {
  const min = Math.max(0, vista.byteLength - 65557);
  for (let i = vista.byteLength - 22; i >= min; i--) {
    if (leerU32(vista, i) === 0x06054b50) return i;
  }
  throw new Error("El archivo no parece un Excel válido (no se encontró el índice del ZIP)");
}

async function inflar(bytes, metodo) {
  if (metodo === 0) return bytes;                    // guardado sin comprimir
  if (metodo !== 8) throw new Error("Compresión no soportada dentro del Excel (método " + metodo + ")");
  if (typeof DecompressionStream === "undefined") {
    throw new Error("Este navegador no puede descomprimir el Excel. Usa Chrome, o exporta el archivo como CSV.");
  }
  const ds = new DecompressionStream("deflate-raw");
  const flujo = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(flujo).arrayBuffer());
}

// Devuelve un Map: nombre de archivo interno → texto
async function abrirZip(buffer, queNecesito) {
  const vista = new DataView(buffer);
  const bytes = new Uint8Array(buffer);
  const eocd = buscarEOCD(vista);
  const total = leerU16(vista, eocd + 10);
  let p = leerU32(vista, eocd + 16);

  const salida = new Map();
  const dec = new TextDecoder("utf-8");

  for (let i = 0; i < total; i++) {
    if (leerU32(vista, p) !== 0x02014b50) break;
    const metodo = leerU16(vista, p + 10);
    const tamComprimido = leerU32(vista, p + 20);
    const largoNombre = leerU16(vista, p + 28);
    const largoExtra = leerU16(vista, p + 30);
    const largoComentario = leerU16(vista, p + 32);
    const offsetLocal = leerU32(vista, p + 42);
    const nombre = dec.decode(bytes.subarray(p + 46, p + 46 + largoNombre));

    if (queNecesito(nombre)) {
      // Saltar al encabezado local, cuyos campos de longitud son los válidos
      if (leerU32(vista, offsetLocal) !== 0x04034b50) throw new Error("Excel corrupto");
      const nombreLocal = leerU16(vista, offsetLocal + 26);
      const extraLocal = leerU16(vista, offsetLocal + 28);
      const inicio = offsetLocal + 30 + nombreLocal + extraLocal;
      const crudo = bytes.subarray(inicio, inicio + tamComprimido);
      salida.set(nombre, dec.decode(await inflar(crudo, metodo)));
    }
    p += 46 + largoNombre + largoExtra + largoComentario;
  }
  return salida;
}

// ---------- XML de la hoja ----------

// "BC12" → 54 (índice de columna, base 0)
function columnaAIndice(ref) {
  const letras = (ref.match(/^[A-Z]+/) || ["A"])[0];
  let n = 0;
  for (const c of letras) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

// Convierte el valor crudo de una celda en el texto que verá el usuario.
// Importante para los teléfonos: Excel los guarda como número (3314451432)
// y hay que escribirlos sin ".0" ni notación científica.
function comoTexto(valor) {
  if (valor === "" || valor == null) return "";
  const n = Number(valor);
  if (!Number.isFinite(n)) return String(valor);
  if (Number.isInteger(n)) return n.toFixed(0);
  return String(valor);
}

export async function leerXLSX(buffer) {
  const partes = await abrirZip(buffer, n =>
    n === "xl/sharedStrings.xml" || n === "xl/worksheets/sheet1.xml" || n === "xl/workbook.xml");

  const hojaXML = partes.get("xl/worksheets/sheet1.xml");
  if (!hojaXML) throw new Error("El Excel no tiene una primera hoja legible");

  const dp = new DOMParser();

  // Texto compartido (así guarda Excel las cadenas repetidas)
  const compartidas = [];
  const ssXML = partes.get("xl/sharedStrings.xml");
  if (ssXML) {
    const doc = dp.parseFromString(ssXML, "application/xml");
    for (const si of doc.getElementsByTagName("si")) {
      // Un <si> puede venir partido en varios <t> (por formato); se concatenan
      let texto = "";
      for (const t of si.getElementsByTagName("t")) texto += t.textContent;
      compartidas.push(texto);
    }
  }

  const doc = dp.parseFromString(hojaXML, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("No se pudo leer la hoja del Excel");

  const filas = [];
  for (const fila of doc.getElementsByTagName("row")) {
    const celdas = [];
    for (const c of fila.getElementsByTagName("c")) {
      const idx = columnaAIndice(c.getAttribute("r") || "A");
      const tipo = c.getAttribute("t");
      let texto = "";
      if (tipo === "s") {
        const v = c.getElementsByTagName("v")[0];
        texto = v ? (compartidas[Number(v.textContent)] ?? "") : "";
      } else if (tipo === "inlineStr") {
        for (const t of c.getElementsByTagName("t")) texto += t.textContent;
      } else {
        const v = c.getElementsByTagName("v")[0];
        texto = v ? comoTexto(v.textContent) : "";
      }
      while (celdas.length < idx) celdas.push(""); // rellena columnas vacías
      celdas[idx] = texto.trim();
    }
    if (celdas.some(x => x !== "")) filas.push(celdas);
  }
  return filas;
}

// ¿Este archivo es un .xlsx? Los ZIP empiezan con "PK".
export function pareceExcel(nombre, primerosBytes) {
  if (/\.xlsx$/i.test(nombre)) return true;
  return primerosBytes && primerosBytes[0] === 0x50 && primerosBytes[1] === 0x4b;
}
