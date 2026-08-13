// ============================================================
// APP PRINCIPAL — Base de Alumnos PRIOSAN
// ============================================================

import * as M from "./model.js";
import * as X from "./export.js";
import { GraphStorage, FolderStorage, FileStorage } from "./storage.js";

// ---------------- estado global ----------------
let db = null;          // la base de datos en memoria
let storage = null;     // proveedor de guardado activo
let dirty = false;      // ¿hay cambios sin guardar?
let vista = { nombre: "buscar", params: {} };
let textoBusqueda = ""; // búsqueda global (vista Buscar)
let destinoExport = "drive"; // "drive" = a la carpeta de Drive · "descarga" = a Descargas

// ¿Podemos escribir en la carpeta de Google Drive conectada?
const puedeGuardarEnDrive = () => !!(storage && typeof storage.guardarExport === "function");

// Entrega una exportación al destino elegido.
async function entregarExport(archivo, etiqueta) {
  if (destinoExport === "drive" && puedeGuardarEnDrive()) {
    try {
      const ruta = await storage.guardarExport(archivo.nombre, archivo.contenido);
      toast(`📄 ${etiqueta} guardado en Google Drive → ${ruta}`);
      return;
    } catch (e) {
      toast("No se pudo guardar en Drive (" + e.message + "). Se descargó a tu computadora.", "error");
    }
  }
  X.descargar(archivo);
  toast(`📄 ${etiqueta} descargado: ${archivo.nombre}`);
}

const $ = sel => document.querySelector(sel);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

// ---------------- toast ----------------
let toastTimer = null;
function toast(msg, tipo = "") {
  const t = $("#toast");
  t.textContent = msg;
  t.className = "toast" + (tipo ? " toast-" + tipo : "");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add("oculto"), tipo === "error" ? 8000 : 4000);
}

// ---------------- cambios sin guardar ----------------
function marcar() {
  dirty = true;
  actualizarChip();
}
function actualizarChip() {
  const chip = $("#indicador-cambios");
  if (dirty) { chip.textContent = "● Cambios sin guardar"; chip.className = "chip chip-alerta"; }
  else { chip.textContent = "Sin cambios"; chip.className = "chip chip-ok"; }
}
window.addEventListener("beforeunload", e => {
  if (dirty) { e.preventDefault(); e.returnValue = ""; }
});

// ============================================================
// ARRANQUE: elegir modo de datos
// ============================================================

async function entrar(datos) {
  db = M.normalizarBase(datos);
  $("#pantalla-inicio").classList.add("oculto");
  if (db.meta.passwordHash) {
    $("#pantalla-candado").classList.remove("oculto");
    $("#input-candado").focus();
  } else {
    mostrarApp();
  }
}

function mostrarApp() {
  $("#pantalla-candado").classList.add("oculto");
  $("#app").classList.remove("oculto");
  actualizarChip();
  navegar("buscar");
}

$("#form-candado").addEventListener("submit", async e => {
  e.preventDefault();
  const h = await M.hashContrasena($("#input-candado").value);
  if (h === db.meta.passwordHash) mostrarApp();
  else $("#candado-error").classList.remove("oculto");
});

$("#btn-abrir-onedrive").addEventListener("click", async () => {
  if (!GraphStorage.configurado()) {
    const aviso = $("#aviso-config");
    aviso.innerHTML = `⚙️ <strong>Falta un paso de configuración (una sola vez):</strong> registrar la app
      con Microsoft para obtener el <em>ID de cliente</em> y escribirlo en <code>js/config.js</code>.
      Las instrucciones están en <strong>GUIA-CONFIGURACION.md</strong>.
      <br><br>Al registrarla, la <strong>dirección de retorno (Redirect URI)</strong> que debes pegar es exactamente:
      <br><code style="user-select:all;background:#fff;padding:3px 6px;border-radius:4px;display:inline-block;margin-top:4px;word-break:break-all">${esc(GraphStorage.redirectUri())}</code>
      <br><br>Mientras tanto puedes usar «Abrir carpeta local» o «Cargar archivo».`;
    aviso.classList.remove("oculto");
    return;
  }
  try {
    storage = new GraphStorage();
    const usuario = await storage.conectar();
    toast("Conectado como " + usuario);
    const datos = await storage.cargar();
    if (datos === null) {
      const seguir = confirm(
        "⚠️ No se encontró la base de datos con esta cuenta (" + usuario + ").\n\n" +
        "Se buscó en tu propio OneDrive y también entre los archivos que otras cuentas hayan compartido contigo, " +
        "en la ruta:\n" + window.APP_CONFIG.dataFilePath + "\n\n" +
        "Si la base YA EXISTE en la cuenta de tu compañero(a), NO continúes: pídele que comparta contigo la carpeta " +
        "de la base con permiso de edición, y vuelve a entrar. Si continúas, crearás una base nueva y vacía, aparte de la de ella/él.\n\n" +
        "¿Crear de todos modos una base nueva y vacía?");
      if (seguir) {
        await entrar(M.baseVacia());
        marcar();
      }
      return;
    }
    if (!M.validarBase(datos)) return toast("El archivo no parece una base de alumnos válida", "error");
    await entrar(datos);
  } catch (e) {
    toast("No se pudo conectar: " + e.message, "error");
  }
});

// El modo OneDrive solo aparece si alguien configuró el ID de Microsoft.
if (GraphStorage.configurado()) $("#btn-abrir-onedrive").classList.remove("oculto");

$("#btn-abrir-carpeta").addEventListener("click", async () => {
  if (!FolderStorage.disponible()) {
    return toast("Este navegador no permite abrir carpetas. Usa «Cargar archivo», o abre la app en Chrome.", "error");
  }
  try {
    storage = new FolderStorage();
    await storage.conectar();
    const datos = await storage.cargar();
    if (datos === null) {
      const seguir = confirm(
        "⚠️ En la carpeta que elegiste no existe «" + window.APP_CONFIG.localFileName + "».\n\n" +
        "Si la base YA EXISTE, seguramente elegiste la carpeta equivocada: cancela y vuelve a intentar " +
        "eligiendo la carpeta compartida de Google Drive.\n\n" +
        "Si continúas, se creará una base nueva y vacía, aparte de la que ya tengan.\n\n" +
        "¿Crear de todos modos una base nueva y vacía aquí?");
      if (seguir) {
        await entrar(M.baseVacia());
        marcar();
      }
      return;
    }
    if (!M.validarBase(datos)) return toast("El archivo no parece una base de alumnos válida", "error");
    await entrar(datos);
  } catch (e) {
    if (e.name !== "AbortError") toast("No se pudo abrir la carpeta: " + e.message, "error");
  }
});

$("#btn-abrir-archivo").addEventListener("click", async () => {
  try {
    storage = new FileStorage();
    const datos = await storage.cargar();
    if (datos === null) return;
    if (!M.validarBase(datos)) return toast("El archivo no parece una base de alumnos válida", "error");
    await entrar(datos);
  } catch (e) {
    toast("No se pudo leer el archivo: " + e.message, "error");
  }
});

$("#btn-base-nueva").addEventListener("click", async () => {
  storage = null; // se elegirá dónde guardar en el primer guardado
  await entrar(M.baseVacia());
  marcar();
  toast("Base nueva creada. Al guardar te preguntaré dónde almacenarla.");
});

// ============================================================
// GUARDAR
// ============================================================

$("#btn-guardar").addEventListener("click", guardar);

async function guardar() {
  if (!db) return;
  if (!storage) return elegirDestinoYGuardar();
  try {
    db.meta.lastSavedAt = new Date().toISOString();
    db.meta.lastSavedBy = storage.tipo === "graph" && storage.cuenta ? storage.cuenta.username : storage.nombre;
    db.meta.rev = (db.meta.rev || 0) + 1;
    await guardarCon(storage);
  } catch (e) {
    toast("⚠️ " + e.message, "error");
  }
}

async function guardarCon(st) {
  await st.guardar(db);
  dirty = false;
  actualizarChip();
  const depurados = st.ultimaLimpieza || 0;
  toast("💾 Guardado (" + st.nombre + ")" +
    (depurados ? ` · ${depurados} respaldo(s) antiguo(s) depurado(s)` : ""));
}

function elegirDestinoYGuardar() {
  const graphOK = GraphStorage.configurado();
  abrirModal(`
    <h3>¿Dónde guardar la base?</h3>
    <div class="opciones-inicio">
      ${graphOK ? `<button class="opcion-inicio" data-dest="graph"><span class="opcion-icono">☁️</span><span class="opcion-texto"><strong>OneDrive (internet)</strong><small>Recomendado</small></span></button>` : ""}
      ${FolderStorage.disponible() ? `<button class="opcion-inicio" data-dest="carpeta"><span class="opcion-icono">📁</span><span class="opcion-texto"><strong>Carpeta de Google Drive</strong><small>Elegir la carpeta sincronizada en esta computadora</small></span></button>` : ""}
      <button class="opcion-inicio" data-dest="archivo"><span class="opcion-icono">📄</span><span class="opcion-texto"><strong>Descargar archivo</strong><small>Guardas tú el archivo donde quieras</small></span></button>
    </div>
    <div class="fila-botones"><button class="btn btn-suave" data-cancelar>Cancelar</button></div>
  `);
  $("#modal-caja").querySelectorAll("[data-dest]").forEach(b => b.addEventListener("click", async () => {
    cerrarModal();
    try {
      let st;
      if (b.dataset.dest === "graph") {
        st = new GraphStorage();
        await st.conectar();
        const existente = await st.cargar();
        if (existente && !confirm("⚠️ Ya existe un archivo de datos en OneDrive.\n¿Reemplazarlo con esta base nueva? (se perdería lo que contiene)")) return;
      } else if (b.dataset.dest === "carpeta") {
        st = new FolderStorage();
        await st.conectar();
        const existente = await st.cargar();
        if (existente && !confirm("⚠️ En esa carpeta ya existe un archivo de datos.\n¿Reemplazarlo con esta base nueva? (se perdería lo que contiene)")) return;
      } else {
        st = new FileStorage();
      }
      storage = st;
      db.meta.lastSavedAt = new Date().toISOString();
      db.meta.rev = (db.meta.rev || 0) + 1;
      await guardarCon(st);
    } catch (e) {
      if (e.name !== "AbortError") toast("No se pudo guardar: " + e.message, "error");
    }
  }));
}

// ============================================================
// NAVEGACIÓN
// ============================================================

document.querySelectorAll(".nav-btn").forEach(b =>
  b.addEventListener("click", () => navegar(b.dataset.vista)));

function navegar(nombre, params = {}) {
  vista = { nombre, params };
  document.querySelectorAll(".nav-btn").forEach(b =>
    b.classList.toggle("activo", b.dataset.vista === nombre ||
      (nombre.startsWith("alumno") && b.dataset.vista === "alumnos") ||
      ((nombre === "curso" || nombre === "edicion") && b.dataset.vista === "cursos")));
  render();
}

function render() {
  const vistas = {
    buscar: renderBuscar,
    alumnos: renderAlumnos,
    alumno: renderAlumno,
    cursos: renderCursos,
    curso: renderCurso,
    edicion: renderEdicion,
    promos: renderPromos,
    exportar: renderExportar
  };
  (vistas[vista.nombre] || renderBuscar)();
}

// ============================================================
// VISTA: BUSCAR
// ============================================================

function renderBuscar() {
  $("#contenido").innerHTML = `
    <div class="tarjeta">
      <input type="text" id="busqueda-global" class="buscador-grande"
        placeholder="Buscar alumno (nombre, apellidos, email) o curso…" value="${esc(textoBusqueda)}">
    </div>
    <div id="resultados-busqueda"></div>
  `;
  const input = $("#busqueda-global");
  input.addEventListener("input", () => { textoBusqueda = input.value; pintarResultados(); });
  input.focus();
  input.setSelectionRange(input.value.length, input.value.length);
  pintarResultados();
}

function pintarResultados() {
  const cont = $("#resultados-busqueda");
  if (!textoBusqueda.trim()) {
    cont.innerHTML = `<div class="tarjeta"><p class="sin-resultados">
      ${db.alumnos.length} alumnos · ${db.cursos.length} cursos · ${db.inscripciones.length} inscripciones registradas.<br>
      Escribe arriba para buscar, o usa los apartados del menú.</p></div>`;
    return;
  }
  const alumnos = M.buscarAlumnos(db, textoBusqueda).slice(0, 30);
  const cursos = M.buscarCursos(db, textoBusqueda).slice(0, 15);
  let html = "";
  if (alumnos.length) {
    html += `<div class="tarjeta"><h2>👥 Alumnos (${alumnos.length})</h2><ul class="lista">` +
      alumnos.map(a => `
        <li data-alumno="${a.id}">
          <span><span class="principal">${esc(M.nombreLista(a))}</span>
            ${a.revisar ? '<span class="etiqueta etiqueta-revisar">revisar</span>' : ""}
            <br><span class="secundario">${esc(a.email || "sin email")} · ${esc(a.telefono || "sin teléfono")}</span></span>
          <span class="derecha">${M.inscripcionesDeAlumno(db, a.id).length} curso(s)</span>
        </li>`).join("") + "</ul></div>";
  }
  if (cursos.length) {
    html += `<div class="tarjeta"><h2>📚 Cursos (${cursos.length})</h2><ul class="lista">` +
      cursos.map(c => `
        <li data-curso="${c.id}">
          <span class="principal">${esc(c.nombre)} ${etiquetaTipo(c)}</span>
          <span class="derecha">${M.edicionesDeCurso(db, c.id).length} edición(es)</span>
        </li>`).join("") + "</ul></div>";
  }
  cont.innerHTML = html || `<div class="tarjeta"><p class="sin-resultados">Sin resultados para «${esc(textoBusqueda)}»</p></div>`;
  cont.querySelectorAll("[data-alumno]").forEach(li => li.addEventListener("click", () => navegar("alumno", { id: li.dataset.alumno })));
  cont.querySelectorAll("[data-curso]").forEach(li => li.addEventListener("click", () => navegar("curso", { id: li.dataset.curso })));
}

function etiquetaTipo(c) {
  return c.tipo === "video"
    ? '<span class="etiqueta etiqueta-video">🎬 video</span>'
    : '<span class="etiqueta etiqueta-vivo">🔴 en vivo</span>';
}

// ============================================================
// VISTA: ALUMNOS (lista alfabética)
// ============================================================

let filtroAlumnos = "";

function renderAlumnos() {
  $("#contenido").innerHTML = `
    <div class="encabezado-vista">
      <h2>👥 Alumnos <span class="secundario" style="font-size:15px;color:var(--color-texto-suave)">(${db.alumnos.length})</span></h2>
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        <input type="text" id="filtro-alumnos" class="buscador" placeholder="Filtrar…" value="${esc(filtroAlumnos)}">
        <button class="btn btn-verde" id="btn-nuevo-alumno">➕ Nuevo alumno</button>
      </div>
    </div>
    <div class="tarjeta"><ul class="lista" id="lista-alumnos"></ul></div>
  `;
  $("#btn-nuevo-alumno").addEventListener("click", () => modalAlumno(null));
  const inp = $("#filtro-alumnos");
  inp.addEventListener("input", () => { filtroAlumnos = inp.value; pintarListaAlumnos(); });
  pintarListaAlumnos();
}

function pintarListaAlumnos() {
  const ul = $("#lista-alumnos");
  let alumnos = filtroAlumnos.trim() ? M.buscarAlumnos(db, filtroAlumnos) : [...db.alumnos];
  alumnos.sort((a, b) => M.nombreLista(a).localeCompare(M.nombreLista(b), "es"));
  if (!alumnos.length) {
    ul.innerHTML = `<p class="sin-resultados">${db.alumnos.length ? "Sin coincidencias" : "Todavía no hay alumnos. Usa «Nuevo alumno» para empezar."}</p>`;
    return;
  }
  let letra = "";
  let html = "";
  for (const a of alumnos) {
    const inicial = (M.nombreLista(a)[0] || "#").toUpperCase();
    if (inicial !== letra) { letra = inicial; html += `<li class="separador-letra">${esc(letra)}</li>`; }
    html += `
      <li data-id="${a.id}">
        <span><span class="principal">${esc(M.nombreLista(a))}</span>
          ${a.revisar ? '<span class="etiqueta etiqueta-revisar">revisar</span>' : ""}
          <br><span class="secundario">${esc(a.email || "sin email")} · ${esc(a.telefono || "sin teléfono")}</span></span>
        <span class="derecha">${M.inscripcionesDeAlumno(db, a.id).length} curso(s)</span>
      </li>`;
  }
  ul.innerHTML = html;
  ul.querySelectorAll("[data-id]").forEach(li => li.addEventListener("click", () => navegar("alumno", { id: li.dataset.id })));
}

// ============================================================
// VISTA: DETALLE DE ALUMNO
// ============================================================

function renderAlumno() {
  const a = db.alumnos.find(x => x.id === vista.params.id);
  if (!a) return navegar("alumnos");
  const inscripciones = M.inscripcionesDeAlumno(db, a.id)
    .sort((x, y) => (y.fecha || "").localeCompare(x.fecha || ""));
  $("#contenido").innerHTML = `
    <div class="migas"><a data-volver>← Alumnos</a></div>
    <div class="tarjeta">
      <div class="encabezado-vista">
        <h2>${esc(M.nombreCompleto(a)) || "(sin nombre)"} ${a.revisar ? '<span class="etiqueta etiqueta-revisar">revisar datos</span>' : ""}</h2>
        <div style="display:flex;gap:8px">
          <button class="btn btn-suave btn-mini" id="btn-editar-alumno">✏️ Editar</button>
          <button class="btn btn-peligro btn-mini" id="btn-eliminar-alumno">🗑 Eliminar</button>
        </div>
      </div>
      <div class="datos-grid">
        <div class="dato"><label>Email</label><div>${esc(a.email || "—")}</div></div>
        <div class="dato"><label>Teléfono</label><div>${esc(a.telefono || "—")}</div></div>
        <div class="dato"><label>Nota</label><div>${esc(a.nota || "—")}</div></div>
      </div>
    </div>
    <div class="tarjeta">
      <div class="encabezado-vista">
        <h2 style="font-size:18px">📚 Historial de cursos (${inscripciones.length})</h2>
        <button class="btn btn-verde btn-mini" id="btn-inscribir">➕ Inscribir a un curso</button>
      </div>
      ${inscripciones.length ? `<table><thead><tr><th>Curso · Edición</th><th>Fecha</th><th></th><th></th></tr></thead><tbody>
        ${inscripciones.map(i => `
          <tr>
            <td>${esc(M.etiquetaEdicion(db, i.edicionId))}${i.nota && !i.promo ? `<br><span class="secundario" style="font-size:12px;color:var(--color-texto-suave)">${esc(i.nota)}</span>` : ""}</td>
            <td>${esc(i.fecha || "")}</td>
            <td>${i.promo ? '<span class="etiqueta etiqueta-promo">🎁 promo</span>' : ""}</td>
            <td style="text-align:right"><button class="btn btn-peligro btn-mini" data-quitar="${i.id}">Quitar</button></td>
          </tr>`).join("")}
      </tbody></table>` : '<p class="sin-resultados">Sin inscripciones todavía</p>'}
    </div>
  `;
  $("[data-volver]").addEventListener("click", () => navegar("alumnos"));
  $("#btn-editar-alumno").addEventListener("click", () => modalAlumno(a));
  $("#btn-eliminar-alumno").addEventListener("click", () => {
    if (confirm(`¿Eliminar a ${M.nombreCompleto(a)} y todas sus inscripciones?\nEsta acción no se puede deshacer.`)) {
      M.eliminarAlumno(db, a.id);
      marcar();
      toast("Alumno eliminado");
      navegar("alumnos");
    }
  });
  $("#btn-inscribir").addEventListener("click", () => modalInscribir({ alumnoId: a.id }));
  document.querySelectorAll("[data-quitar]").forEach(b => b.addEventListener("click", () => {
    if (confirm("¿Quitar esta inscripción? (si generó un curso de regalo, también se quita)")) {
      M.eliminarInscripcion(db, b.dataset.quitar);
      marcar();
      render();
    }
  }));
}

// ============================================================
// VISTA: CURSOS
// ============================================================

function renderCursos() {
  const cursos = [...db.cursos].sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  $("#contenido").innerHTML = `
    <div class="encabezado-vista">
      <h2>📚 Cursos <span style="font-size:15px;color:var(--color-texto-suave)">(${cursos.length})</span></h2>
      <button class="btn btn-verde" id="btn-nuevo-curso">➕ Nuevo curso</button>
    </div>
    <div class="tarjeta">
      ${cursos.length ? `<ul class="lista">${cursos.map(c => {
        const eds = M.edicionesDeCurso(db, c.id);
        const total = eds.reduce((n, e) => n + M.inscritosDeEdicion(db, e.id).length, 0);
        return `<li data-id="${c.id}">
          <span><span class="principal">${esc(c.nombre)}</span> ${etiquetaTipo(c)}
            ${c.descripcion ? `<br><span class="secundario">${esc(c.descripcion)}</span>` : ""}</span>
          <span class="derecha">${eds.length} edición(es)<br>${total} inscripción(es)</span>
        </li>`;
      }).join("")}</ul>` : '<p class="sin-resultados">Todavía no hay cursos registrados</p>'}
    </div>
  `;
  $("#btn-nuevo-curso").addEventListener("click", () => modalCurso(null));
  document.querySelectorAll("[data-id]").forEach(li => li.addEventListener("click", () => navegar("curso", { id: li.dataset.id })));
}

// ============================================================
// VISTA: DETALLE DE CURSO
// ============================================================

function renderCurso() {
  const c = db.cursos.find(x => x.id === vista.params.id);
  if (!c) return navegar("cursos");
  const eds = M.edicionesDeCurso(db, c.id);
  $("#contenido").innerHTML = `
    <div class="migas"><a data-volver>← Cursos</a></div>
    <div class="tarjeta">
      <div class="encabezado-vista">
        <h2>${esc(c.nombre)} ${etiquetaTipo(c)}</h2>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="btn btn-suave btn-mini" id="btn-editar-curso">✏️ Editar</button>
          <button class="btn btn-suave btn-mini" id="btn-exportar-curso">📤 Exportar todo</button>
          <button class="btn btn-peligro btn-mini" id="btn-eliminar-curso">🗑 Eliminar</button>
        </div>
      </div>
      ${c.descripcion ? `<p style="color:var(--color-texto-suave)">${esc(c.descripcion)}</p>` : ""}
    </div>
    <div class="tarjeta">
      <div class="encabezado-vista">
        <h2 style="font-size:18px">🗓 Ediciones (${eds.length})</h2>
        <button class="btn btn-verde btn-mini" id="btn-nueva-edicion">➕ Nueva edición</button>
      </div>
      ${eds.length ? `<ul class="lista">${eds.map(e => `
        <li data-ed="${e.id}">
          <span class="principal">${esc(e.periodo || "(sin fecha)")}${e.nota ? ` <span class="secundario">· ${esc(e.nota)}</span>` : ""}</span>
          <span class="derecha">${M.inscritosDeEdicion(db, e.id).length} inscrito(s)</span>
        </li>`).join("")}</ul>` : '<p class="sin-resultados">Sin ediciones. Agrega la primera con «Nueva edición».</p>'}
    </div>
  `;
  $("[data-volver]").addEventListener("click", () => navegar("cursos"));
  $("#btn-editar-curso").addEventListener("click", () => modalCurso(c));
  $("#btn-exportar-curso").addEventListener("click", () => entregarExport(X.construirCurso(db, c.id), "Curso «" + c.nombre + "»"));
  $("#btn-eliminar-curso").addEventListener("click", () => {
    if (confirm(`¿Eliminar el curso «${c.nombre}» con TODAS sus ediciones e inscripciones?\nEsta acción no se puede deshacer.`)) {
      M.eliminarCurso(db, c.id);
      marcar();
      navegar("cursos");
    }
  });
  $("#btn-nueva-edicion").addEventListener("click", () => modalEdicion(c.id, null));
  document.querySelectorAll("[data-ed]").forEach(li => li.addEventListener("click", () => navegar("edicion", { id: li.dataset.ed })));
}

// ============================================================
// VISTA: DETALLE DE EDICIÓN (lista de inscritos)
// ============================================================

function renderEdicion() {
  const e = db.ediciones.find(x => x.id === vista.params.id);
  if (!e) return navegar("cursos");
  const c = db.cursos.find(x => x.id === e.cursoId);
  const inscritos = M.inscritosDeEdicion(db, e.id)
    .map(i => ({ i, a: db.alumnos.find(x => x.id === i.alumnoId) }))
    .sort((x, y) => M.nombreLista(x.a || {}).localeCompare(M.nombreLista(y.a || {}), "es"));
  // Quién no podría recibir su constancia por correo (no bloquea nada: es un recordatorio).
  const sinEmail = inscritos.filter(({ a }) => a && !(a.email || "").trim()).map(({ a }) => a);
  $("#contenido").innerHTML = `
    <div class="migas"><a data-volver-cursos>← Cursos</a> / <a data-volver-curso>${esc(c?.nombre || "?")}</a></div>
    <div class="tarjeta">
      <div class="encabezado-vista">
        <h2>${esc(c?.nombre || "?")} · ${esc(e.periodo || "sin fecha")}</h2>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="btn btn-verde btn-mini" id="btn-inscribir-aqui">➕ Inscribir alumno</button>
          <button class="btn btn-suave btn-mini" id="btn-constancias">🎓 Constancias</button>
          <button class="btn btn-suave btn-mini" id="btn-copiar-nombres">📋 Copiar nombres</button>
          <button class="btn btn-suave btn-mini" id="btn-editar-edicion">✏️ Editar</button>
          <button class="btn btn-suave btn-mini" id="btn-exportar-edicion">📤 Exportar</button>
          <button class="btn btn-peligro btn-mini" id="btn-eliminar-edicion">🗑 Eliminar</button>
        </div>
      </div>
      ${e.nota ? `<p style="color:var(--color-texto-suave)">${esc(e.nota)}</p>` : ""}
      ${sinEmail.length ? `<div class="aviso-duplicado" style="margin-top:12px">
        📧 <strong>${sinEmail.length} sin email:</strong>
        ${sinEmail.map(a => `<a data-ir-alumno="${a.id}">${esc(M.nombreCompleto(a))}</a>`).join(" · ")}
        <br><span style="color:var(--color-texto-suave);font-size:13px">
        El archivo de constancias se genera igual, con esa casilla en blanco; solo que
        no podrán recibirla por correo. Haz clic en el nombre para capturar su email.</span>
      </div>` : ""}
      <h3 style="margin:10px 0 6px;font-size:16px">Inscritos (${inscritos.length})</h3>
      ${inscritos.length ? `<table><thead><tr><th>Alumno</th><th>Email</th><th>Fecha</th><th></th><th></th></tr></thead><tbody>
        ${inscritos.map(({ i, a }) => `
          <tr class="clicable" data-alumno="${a?.id || ""}">
            <td>${esc(a ? M.nombreLista(a) : "(alumno eliminado)")}</td>
            <td>${esc(a?.email || "")}</td>
            <td>${esc(i.fecha || "")}</td>
            <td>${i.promo ? '<span class="etiqueta etiqueta-promo">🎁 promo</span>' : ""}</td>
            <td style="text-align:right"><button class="btn btn-peligro btn-mini" data-quitar="${i.id}">Quitar</button></td>
          </tr>`).join("")}
      </tbody></table>` : '<p class="sin-resultados">Nadie inscrito todavía</p>'}
    </div>
  `;
  $("[data-volver-cursos]").addEventListener("click", () => navegar("cursos"));
  $("[data-volver-curso]").addEventListener("click", () => navegar("curso", { id: e.cursoId }));
  document.querySelectorAll("[data-ir-alumno]").forEach(a =>
    a.addEventListener("click", () => navegar("alumno", { id: a.dataset.irAlumno })));
  $("#btn-inscribir-aqui").addEventListener("click", () => modalInscribir({ edicionId: e.id }));
  $("#btn-editar-edicion").addEventListener("click", () => modalEdicion(e.cursoId, e));
  $("#btn-exportar-edicion").addEventListener("click", () => entregarExport(X.construirEdicion(db, e.id), "Lista de inscritos"));
  $("#btn-constancias").addEventListener("click", () => {
    if (!inscritos.length) return toast("Esta edición no tiene inscritos todavía", "error");
    const archivo = X.construirConstancias(db, e.id);
    entregarExport(archivo, `Constancias (${archivo.total} alumnos)`);
  });
  $("#btn-copiar-nombres").addEventListener("click", async () => {
    const texto = X.nombresParaPegar(db, e.id);
    if (!texto) return toast("Esta edición no tiene inscritos todavía", "error");
    const n = texto.split("\n").length;
    try {
      await navigator.clipboard.writeText(texto);
      toast(`📋 ${n} nombre(s) copiados — ya puedes pegarlos en Canva`);
    } catch {
      // Si el navegador bloquea el portapapeles, mostramos la lista para copiar a mano.
      abrirModal(`
        <h3>📋 Nombres para copiar (${n})</h3>
        <p style="color:var(--color-texto-suave);font-size:13.5px;margin-bottom:10px">
          Selecciona todo el texto y cópialo con Cmd+C.</p>
        <textarea readonly style="width:100%;height:260px;font-family:inherit;padding:10px;
          border:1px solid var(--color-borde);border-radius:8px">${esc(texto)}</textarea>
        <div class="fila-botones"><button class="btn btn-suave" data-cancelar>Cerrar</button></div>
      `);
      $("#modal-caja").querySelector("textarea").select();
    }
  });
  $("#btn-eliminar-edicion").addEventListener("click", () => {
    if (confirm("¿Eliminar esta edición y todas sus inscripciones?")) {
      const cursoId = e.cursoId;
      M.eliminarEdicion(db, e.id);
      marcar();
      navegar("curso", { id: cursoId });
    }
  });
  document.querySelectorAll("[data-alumno]").forEach(tr => tr.addEventListener("click", ev => {
    if (ev.target.closest("[data-quitar]")) return;
    if (tr.dataset.alumno) navegar("alumno", { id: tr.dataset.alumno });
  }));
  document.querySelectorAll("[data-quitar]").forEach(b => b.addEventListener("click", () => {
    if (confirm("¿Quitar esta inscripción?")) {
      M.eliminarInscripcion(db, b.dataset.quitar);
      marcar();
      render();
    }
  }));
}

// ============================================================
// VISTA: PROMOCIONES
// ============================================================

function renderPromos() {
  const promos = db.promociones;
  $("#contenido").innerHTML = `
    <div class="encabezado-vista">
      <h2>🎁 Promociones</h2>
      <button class="btn btn-verde" id="btn-nueva-promo">➕ Nueva regla</button>
    </div>
    <div class="tarjeta">
      <p style="color:var(--color-texto-suave);margin-bottom:12px">
        Al inscribir a un alumno en el curso disparador, la app agrega sola la inscripción al curso de regalo.
      </p>
      ${promos.length ? `<ul class="lista">${promos.map(p => {
        const trigger = db.cursos.find(c => c.id === p.triggerCursoId);
        return `<li style="cursor:default">
          <span class="principal" style="${p.activa ? "" : "opacity:.45;text-decoration:line-through"}">
            ${esc(trigger?.nombre || "(curso eliminado)")} → 🎁 ${esc(M.etiquetaEdicion(db, p.giftEdicionId))}
          </span>
          <span style="display:flex;gap:8px">
            <button class="btn btn-suave btn-mini" data-toggle="${p.id}">${p.activa ? "Pausar" : "Activar"}</button>
            <button class="btn btn-peligro btn-mini" data-borrar="${p.id}">Eliminar</button>
          </span>
        </li>`;
      }).join("")}</ul>` : '<p class="sin-resultados">No hay reglas de promoción</p>'}
    </div>
  `;
  $("#btn-nueva-promo").addEventListener("click", modalPromo);
  document.querySelectorAll("[data-toggle]").forEach(b => b.addEventListener("click", () => {
    const p = db.promociones.find(x => x.id === b.dataset.toggle);
    p.activa = !p.activa;
    marcar();
    render();
  }));
  document.querySelectorAll("[data-borrar]").forEach(b => b.addEventListener("click", () => {
    if (confirm("¿Eliminar esta regla de promoción? (las inscripciones ya creadas no se tocan)")) {
      M.eliminarPromocion(db, b.dataset.borrar);
      marcar();
      render();
    }
  }));
}

// ============================================================
// VISTA: EXPORTAR / SEGURIDAD
// ============================================================

function renderExportar() {
  // Si no hay carpeta de Drive conectada, el destino real es la descarga.
  const dest = puedeGuardarEnDrive() ? destinoExport : "descarga";
  $("#contenido").innerHTML = `
    <div class="encabezado-vista"><h2>📤 Exportar y respaldos</h2></div>
    <div class="tarjeta">
      <h3 style="margin-bottom:10px;font-size:17px">¿Dónde quieres los archivos?</h3>
      <div class="fila-botones" style="justify-content:flex-start">
        <button class="btn ${dest === "drive" ? "btn-verde" : "btn-suave"}" id="dest-drive"
          ${puedeGuardarEnDrive() ? "" : "disabled style='opacity:.5;cursor:not-allowed' title='Disponible al abrir la base desde la carpeta de Google Drive'"}>
          ${dest === "drive" ? "✓ " : ""}📁 Guardar en Google Drive
        </button>
        <button class="btn ${dest === "descarga" ? "btn-verde" : "btn-suave"}" id="dest-descarga">
          ${dest === "descarga" ? "✓ " : ""}⬇️ Descargar a esta computadora
        </button>
      </div>
      <p style="color:var(--color-texto-suave);font-size:13.5px;margin-top:10px">
        ${puedeGuardarEnDrive()
          ? "Guardar en Drive es lo más cómodo: los archivos aparecen en la subcarpeta «exportaciones» y desde Drive se abren con Google Sheets de un clic, igual en la Mac que en la Chromebook."
          : "Para guardar directo en Drive, abre la base con «Abrir carpeta de Google Drive» desde la pantalla de inicio."}
      </p>
    </div>
    <div class="tarjeta">
      <h3 style="margin-bottom:10px;font-size:17px">Exportar a hoja de cálculo</h3>
      <div class="fila-botones" style="justify-content:flex-start">
        <button class="btn btn-verde" id="exp-alumnos">👥 Todos los alumnos</button>
        <button class="btn btn-verde" id="exp-inscripciones">📋 Todas las inscripciones</button>
      </div>
      <p style="color:var(--color-texto-suave);font-size:13.5px;margin-top:10px">
        Para exportar un curso o edición específica, entra al curso en el apartado «Cursos» y usa su botón «Exportar».
        <br><strong>Para abrirlos:</strong> en Google Drive, clic derecho sobre el archivo →
        «Abrir con» → «Google Sheets». Funciona igual en Mac y en Chromebook, sin instalar nada.
      </p>
    </div>
    <div class="tarjeta">
      <h3 style="margin-bottom:10px;font-size:17px">Respaldo completo</h3>
      <div class="fila-botones" style="justify-content:flex-start">
        <button class="btn btn-suave" id="exp-json">💾 Descargar respaldo (JSON)</button>
      </div>
      <p style="color:var(--color-texto-suave);font-size:13.5px;margin-top:10px">
        Es el archivo completo de la base. Se puede volver a cargar con «Cargar archivo» en la pantalla de inicio.
        Además, cada vez que guardas, la app deja automáticamente una copia con fecha en la subcarpeta «respaldos»
        de la misma carpeta de Google Drive.
      </p>
    </div>
    <div class="tarjeta">
      <h3 style="margin-bottom:10px;font-size:17px">🔒 Contraseña de la app</h3>
      <p style="color:var(--color-texto-suave);font-size:13.5px;margin-bottom:10px">
        Candado al abrir la base. Es un disuasivo, no una protección total: la protección real
        son los permisos de compartición de Google Drive.
        ${db.meta.passwordHash ? "<strong>Contraseña activa.</strong>" : "<strong>Sin contraseña.</strong>"}
      </p>
      <div class="fila-botones" style="justify-content:flex-start">
        <button class="btn btn-suave" id="btn-cambiar-pass">${db.meta.passwordHash ? "Cambiar contraseña" : "Poner contraseña"}</button>
        ${db.meta.passwordHash ? '<button class="btn btn-peligro" id="btn-quitar-pass">Quitar contraseña</button>' : ""}
      </div>
    </div>
    <div class="tarjeta">
      <h3 style="margin-bottom:10px;font-size:17px">ℹ️ Estado</h3>
      <div class="datos-grid">
        <div class="dato"><label>Almacenamiento</label><div>${esc(storage ? storage.nombre : "Sin elegir (se elige al guardar)")}</div></div>
        <div class="dato"><label>Último guardado</label><div>${db.meta.lastSavedAt ? new Date(db.meta.lastSavedAt).toLocaleString("es-MX") : "—"}</div></div>
        <div class="dato"><label>Guardado por</label><div>${esc(db.meta.lastSavedBy || "—")}</div></div>
      </div>
    </div>
  `;
  $("#dest-drive").addEventListener("click", () => { destinoExport = "drive"; render(); });
  $("#dest-descarga").addEventListener("click", () => { destinoExport = "descarga"; render(); });
  $("#exp-alumnos").addEventListener("click", () => entregarExport(X.construirAlumnos(db), "Listado de alumnos"));
  $("#exp-inscripciones").addEventListener("click", () => entregarExport(X.construirInscripciones(db), "Listado de inscripciones"));
  $("#exp-json").addEventListener("click", () => entregarExport(X.construirRespaldoJSON(db), "Respaldo completo"));
  $("#btn-cambiar-pass").addEventListener("click", async () => {
    const p1 = prompt("Nueva contraseña (mínimo 4 caracteres):");
    if (!p1) return;
    if (p1.length < 4) return toast("Muy corta (mínimo 4 caracteres)", "error");
    const p2 = prompt("Repite la contraseña:");
    if (p1 !== p2) return toast("No coinciden", "error");
    db.meta.passwordHash = await M.hashContrasena(p1);
    marcar();
    toast("🔒 Contraseña establecida (recuerda Guardar)");
    render();
  });
  const q = $("#btn-quitar-pass");
  if (q) q.addEventListener("click", () => {
    if (confirm("¿Quitar la contraseña de entrada?")) {
      db.meta.passwordHash = null;
      marcar();
      toast("Contraseña eliminada (recuerda Guardar)");
      render();
    }
  });
}

// ============================================================
// MODALES
// ============================================================

function abrirModal(html) {
  $("#modal-caja").innerHTML = html;
  $("#modal-fondo").classList.remove("oculto");
  const cancelar = $("#modal-caja").querySelector("[data-cancelar]");
  if (cancelar) cancelar.addEventListener("click", cerrarModal);
}
function cerrarModal() {
  $("#modal-fondo").classList.add("oculto");
  $("#modal-caja").innerHTML = "";
}
$("#modal-fondo").addEventListener("click", e => { if (e.target.id === "modal-fondo") cerrarModal(); });

// ---------- modal alumno (nuevo / editar) ----------
function modalAlumno(alumno, alGuardar = null) {
  const esNuevo = !alumno;
  abrirModal(`
    <h3>${esNuevo ? "➕ Nuevo alumno" : "✏️ Editar alumno"}</h3>
    <form id="form-alumno">
      <div class="form-grid">
        <div class="campo ancho-total"><label>Nombre(s) *</label><input name="nombres" required value="${esc(alumno?.nombres)}"></div>
        <div class="campo"><label>Apellido paterno</label><input name="apellidoPaterno" value="${esc(alumno?.apellidoPaterno)}"></div>
        <div class="campo"><label>Apellido materno</label><input name="apellidoMaterno" value="${esc(alumno?.apellidoMaterno)}"></div>
        <div class="campo"><label>Email</label><input name="email" type="email" value="${esc(alumno?.email)}"></div>
        <div class="campo"><label>Teléfono</label><input name="telefono" value="${esc(alumno?.telefono)}"></div>
        <div class="campo ancho-total"><label>Nota</label><textarea name="nota">${esc(alumno?.nota)}</textarea></div>
        <label class="ancho-total" style="display:flex;align-items:center;gap:8px;font-size:14px">
          <input type="checkbox" name="revisar" ${alumno?.revisar ? "checked" : ""}> Marcar para revisión (datos dudosos)
        </label>
      </div>
      <div id="aviso-dup"></div>
      <div class="fila-botones">
        <button type="button" class="btn btn-suave" data-cancelar>Cancelar</button>
        <button type="submit" class="btn btn-primario">${esNuevo ? "Registrar" : "Guardar cambios"}</button>
      </div>
    </form>
  `);
  const form = $("#form-alumno");
  let confirmadoDuplicado = false;

  form.addEventListener("submit", e => {
    e.preventDefault();
    const fd = new FormData(form);
    const datos = {
      nombres: fd.get("nombres"), apellidoPaterno: fd.get("apellidoPaterno"),
      apellidoMaterno: fd.get("apellidoMaterno"), email: fd.get("email"),
      telefono: fd.get("telefono"), nota: fd.get("nota"), revisar: !!fd.get("revisar")
    };
    const dups = M.detectarDuplicados(db, datos, alumno?.id || null);
    if (dups.length && !confirmadoDuplicado) {
      $("#aviso-dup").innerHTML = `<div class="aviso-duplicado">
        ⚠️ <strong>Puede que este alumno ya exista:</strong>
        <ul>${dups.map(d => `<li><a data-usar="${d.alumno.id}">${esc(M.nombreCompleto(d.alumno))}</a>
          <span style="color:var(--color-texto-suave)">(${d.porEmail ? "mismo email" : "mismo nombre"} · ${esc(d.alumno.email || "sin email")})</span></li>`).join("")}</ul>
        <p style="margin-top:6px">Haz clic en el nombre para <strong>usar ese alumno existente</strong>, o vuelve a pulsar
        «${esNuevo ? "Registrar" : "Guardar cambios"}» para ${esNuevo ? "registrarlo de todos modos como persona distinta" : "guardar de todos modos"}.</p>
      </div>`;
      confirmadoDuplicado = true;
      document.querySelectorAll("[data-usar]").forEach(a => a.addEventListener("click", () => {
        cerrarModal();
        if (alGuardar) alGuardar(db.alumnos.find(x => x.id === a.dataset.usar));
        else navegar("alumno", { id: a.dataset.usar });
      }));
      return;
    }
    let resultado;
    if (esNuevo) {
      resultado = M.agregarAlumno(db, datos);
      toast("✅ Alumno registrado");
    } else {
      resultado = M.actualizarAlumno(db, alumno.id, datos);
      toast("✅ Datos actualizados");
    }
    marcar();
    cerrarModal();
    if (alGuardar) alGuardar(resultado);
    else if (esNuevo) navegar("alumno", { id: resultado.id });
    else render();
  });
}

// ---------- modal curso (nuevo / editar) ----------
function modalCurso(curso, alGuardar = null) {
  const esNuevo = !curso;
  abrirModal(`
    <h3>${esNuevo ? "➕ Nuevo curso" : "✏️ Editar curso"}</h3>
    <form id="form-curso">
      <div class="form-grid">
        <div class="campo ancho-total"><label>Nombre del curso *</label><input name="nombre" required value="${esc(curso?.nombre)}"></div>
        <div class="campo"><label>Tipo</label>
          <select name="tipo">
            <option value="vivo" ${curso?.tipo !== "video" ? "selected" : ""}>🔴 En vivo</option>
            <option value="video" ${curso?.tipo === "video" ? "selected" : ""}>🎬 En video / antiguo</option>
          </select>
        </div>
        <div class="campo"><label>Categoría / descripción</label><input name="descripcion" value="${esc(curso?.descripcion)}"></div>
      </div>
      <div id="aviso-dup-curso"></div>
      <div class="fila-botones">
        <button type="button" class="btn btn-suave" data-cancelar>Cancelar</button>
        <button type="submit" class="btn btn-primario">${esNuevo ? "Registrar" : "Guardar cambios"}</button>
      </div>
    </form>
  `);
  const form = $("#form-curso");
  let confirmado = false;
  form.addEventListener("submit", e => {
    e.preventDefault();
    const fd = new FormData(form);
    const nombre = fd.get("nombre");
    const dup = M.cursoDuplicado(db, nombre, curso?.id || null);
    if (dup && !confirmado) {
      $("#aviso-dup-curso").innerHTML = `<div class="aviso-duplicado">
        ⚠️ Ya existe un curso llamado <strong>«${esc(dup.nombre)}»</strong>.
        Si es el mismo curso, cancela y agrégale una <strong>edición nueva</strong> en lugar de duplicarlo.
        Vuelve a pulsar el botón para crearlo de todos modos.</div>`;
      confirmado = true;
      return;
    }
    let c;
    if (esNuevo) {
      c = M.agregarCurso(db, { nombre, tipo: fd.get("tipo"), descripcion: fd.get("descripcion") });
      toast("✅ Curso registrado");
    } else {
      Object.assign(curso, { nombre: nombre.trim(), tipo: fd.get("tipo"), descripcion: (fd.get("descripcion") || "").trim() });
      c = curso;
      toast("✅ Curso actualizado");
    }
    marcar();
    cerrarModal();
    if (alGuardar) alGuardar(c);
    else render();
  });
}

// ---------- modal edición (nueva / editar) ----------
function modalEdicion(cursoId, edicion, alGuardar = null) {
  const curso = db.cursos.find(c => c.id === cursoId);
  const esNueva = !edicion;
  abrirModal(`
    <h3>${esNueva ? "➕ Nueva edición" : "✏️ Editar edición"} · ${esc(curso?.nombre || "")}</h3>
    <form id="form-edicion">
      <div class="form-grid">
        <div class="campo"><label>Fecha o periodo *</label><input name="periodo" required placeholder="Ej. Agosto 2026" value="${esc(edicion?.periodo)}"></div>
        <div class="campo"><label>Nota (opcional)</label><input name="nota" value="${esc(edicion?.nota)}"></div>
      </div>
      <div class="fila-botones">
        <button type="button" class="btn btn-suave" data-cancelar>Cancelar</button>
        <button type="submit" class="btn btn-primario">${esNueva ? "Agregar" : "Guardar cambios"}</button>
      </div>
    </form>
  `);
  $("#form-edicion").addEventListener("submit", e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    let ed;
    if (esNueva) {
      ed = M.agregarEdicion(db, { cursoId, periodo: fd.get("periodo"), nota: fd.get("nota") });
      toast("✅ Edición agregada");
    } else {
      Object.assign(edicion, { periodo: (fd.get("periodo") || "").trim(), nota: (fd.get("nota") || "").trim() });
      ed = edicion;
      toast("✅ Edición actualizada");
    }
    marcar();
    cerrarModal();
    if (alGuardar) alGuardar(ed);
    else render();
  });
}

// ---------- modal promoción ----------
function modalPromo() {
  const cursos = [...db.cursos].sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  if (cursos.length < 2) return toast("Necesitas al menos dos cursos registrados para crear una promoción", "error");
  abrirModal(`
    <h3>🎁 Nueva regla de promoción</h3>
    <form id="form-promo">
      <div class="campo" style="margin-bottom:12px">
        <label>Al inscribirse al curso… (disparador)</label>
        <select name="trigger" required>
          <option value="">— elegir curso —</option>
          ${cursos.map(c => `<option value="${c.id}">${esc(c.nombre)}</option>`).join("")}
        </select>
      </div>
      <div class="campo" style="margin-bottom:12px">
        <label>…se regala el curso:</label>
        <select name="giftCurso" required>
          <option value="">— elegir curso —</option>
          ${cursos.map(c => `<option value="${c.id}">${esc(c.nombre)}</option>`).join("")}
        </select>
      </div>
      <div class="campo" id="campo-gift-edicion" style="margin-bottom:12px;display:none">
        <label>Edición de regalo:</label>
        <select name="giftEdicion"></select>
      </div>
      <div class="fila-botones">
        <button type="button" class="btn btn-suave" data-cancelar>Cancelar</button>
        <button type="submit" class="btn btn-primario">Crear regla</button>
      </div>
    </form>
  `);
  const selCurso = document.querySelector("[name=giftCurso]");
  const campoEd = $("#campo-gift-edicion");
  const selEd = document.querySelector("[name=giftEdicion]");
  selCurso.addEventListener("change", () => {
    const eds = M.edicionesDeCurso(db, selCurso.value);
    if (!selCurso.value) { campoEd.style.display = "none"; return; }
    if (!eds.length) {
      campoEd.style.display = "block";
      selEd.innerHTML = `<option value="">(este curso no tiene ediciones — créala primero)</option>`;
      return;
    }
    campoEd.style.display = "block";
    selEd.innerHTML = eds.map(e => `<option value="${e.id}">${esc(e.periodo || "sin fecha")}</option>`).join("");
  });
  $("#form-promo").addEventListener("submit", e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    if (!fd.get("giftEdicion")) return toast("El curso de regalo necesita al menos una edición. Créala primero en «Cursos».", "error");
    if (fd.get("trigger") === fd.get("giftCurso")) return toast("El curso disparador y el de regalo no pueden ser el mismo", "error");
    M.agregarPromocion(db, { triggerCursoId: fd.get("trigger"), giftEdicionId: fd.get("giftEdicion") });
    marcar();
    cerrarModal();
    toast("🎁 Regla de promoción creada");
    render();
  });
}

// ---------- modal inscribir (el flujo completo) ----------
function modalInscribir({ alumnoId = null, edicionId = null }) {
  const alumnoFijo = alumnoId ? db.alumnos.find(a => a.id === alumnoId) : null;
  const edicionFija = edicionId ? db.ediciones.find(e => e.id === edicionId) : null;
  const cursos = [...db.cursos].sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));
  let alumnoSel = alumnoFijo;

  abrirModal(`
    <h3>📝 Nueva inscripción</h3>
    <form id="form-inscribir">
      <div class="campo" style="margin-bottom:12px">
        <label>Alumno *</label>
        ${alumnoFijo
          ? `<div style="padding:9px 12px;background:#f2f8f6;border-radius:8px">${esc(M.nombreCompleto(alumnoFijo))}</div>`
          : `<input type="text" id="ins-buscar-alumno" placeholder="Escribe nombre o email para buscar…" autocomplete="off">
             <div id="ins-resultados"></div>
             <div id="ins-alumno-sel" class="oculto" style="margin-top:6px"></div>
             <div style="margin-top:6px"><a id="ins-nuevo-alumno" style="color:var(--color-primario);cursor:pointer;font-size:14px;font-weight:600">➕ Registrar alumno nuevo</a></div>`}
      </div>
      <div class="campo" style="margin-bottom:12px">
        <label>Curso *</label>
        ${edicionFija
          ? `<div style="padding:9px 12px;background:#f2f8f6;border-radius:8px">${esc(M.etiquetaEdicion(db, edicionFija.id))}</div>`
          : `<select id="ins-curso" required>
              <option value="">— elegir curso —</option>
              ${cursos.map(c => `<option value="${c.id}">${esc(c.nombre)}${c.tipo === "video" ? " (video)" : ""}</option>`).join("")}
              <option value="__nuevo__">➕ Crear curso nuevo…</option>
            </select>`}
      </div>
      ${edicionFija ? "" : `
      <div class="campo oculto" id="campo-ins-edicion" style="margin-bottom:12px">
        <label>Edición *</label>
        <select id="ins-edicion"></select>
        <input type="text" id="ins-nueva-edicion" class="oculto" placeholder="Fecha o periodo de la edición nueva, ej. Agosto 2026" style="margin-top:6px">
      </div>`}
      <div class="form-grid">
        <div class="campo"><label>Fecha de inscripción</label><input type="date" name="fecha" value="${M.hoyISO()}"></div>
        <div class="campo"><label>Nota (opcional)</label><input name="nota"></div>
      </div>
      <div class="fila-botones">
        <button type="button" class="btn btn-suave" data-cancelar>Cancelar</button>
        <button type="submit" class="btn btn-primario">Inscribir</button>
      </div>
    </form>
  `);

  // --- selección de alumno (si no viene fijo) ---
  if (!alumnoFijo) {
    const inputBuscar = $("#ins-buscar-alumno");
    const divRes = $("#ins-resultados");
    const divSel = $("#ins-alumno-sel");

    const pintarSel = () => {
      if (alumnoSel) {
        inputBuscar.classList.add("oculto");
        divRes.innerHTML = "";
        divSel.classList.remove("oculto");
        divSel.innerHTML = `<div style="padding:9px 12px;background:#e3f2ec;border-radius:8px;display:flex;justify-content:space-between;align-items:center">
          <span>✅ ${esc(M.nombreCompleto(alumnoSel))} <span style="color:var(--color-texto-suave)">${esc(alumnoSel.email || "")}</span></span>
          <button type="button" class="btn btn-suave btn-mini" id="ins-cambiar">Cambiar</button></div>`;
        $("#ins-cambiar").addEventListener("click", () => {
          alumnoSel = null;
          divSel.classList.add("oculto");
          inputBuscar.classList.remove("oculto");
          inputBuscar.focus();
        });
      }
    };

    inputBuscar.addEventListener("input", () => {
      const res = M.buscarAlumnos(db, inputBuscar.value).slice(0, 8);
      divRes.innerHTML = res.length
        ? `<ul class="lista" style="border:1px solid var(--color-borde);border-radius:8px;margin-top:6px">${res.map(a =>
            `<li data-pick="${a.id}"><span><span class="principal">${esc(M.nombreLista(a))}</span><br>
             <span class="secundario">${esc(a.email || "sin email")}</span></span></li>`).join("")}</ul>`
        : (inputBuscar.value.trim() ? `<p style="font-size:13px;color:var(--color-texto-suave);margin-top:6px">Sin coincidencias — puedes registrarlo como nuevo ⤵</p>` : "");
      divRes.querySelectorAll("[data-pick]").forEach(li => li.addEventListener("click", () => {
        alumnoSel = db.alumnos.find(a => a.id === li.dataset.pick);
        pintarSel();
      }));
    });

    $("#ins-nuevo-alumno").addEventListener("click", () => {
      // guarda lo capturado del formulario de inscripción y abre el modal de alumno;
      // al terminar, re-abre este modal con el alumno ya seleccionado
      const fecha = document.querySelector("[name=fecha]").value;
      const nota = document.querySelector("[name=nota]").value;
      const cursoSel = edicionFija ? null : $("#ins-curso")?.value;
      cerrarModal();
      modalAlumno(null, nuevoAlumno => {
        modalInscribir({ alumnoId: nuevoAlumno.id, edicionId });
        const f = document.querySelector("[name=fecha]"); if (f && fecha) f.value = fecha;
        const n = document.querySelector("[name=nota]"); if (n) n.value = nota;
        const c = $("#ins-curso"); if (c && cursoSel && cursoSel !== "__nuevo__") { c.value = cursoSel; c.dispatchEvent(new Event("change")); }
      });
    });
    inputBuscar.focus();
  }

  // --- selección de curso / edición (si no viene fija) ---
  if (!edicionFija) {
    const selCurso = $("#ins-curso");
    const campoEd = $("#campo-ins-edicion");
    const selEd = $("#ins-edicion");
    const inputNuevaEd = $("#ins-nueva-edicion");

    selCurso.addEventListener("change", () => {
      if (selCurso.value === "__nuevo__") {
        // crear curso al vuelo: se abre el mini-formulario y al terminar
        // se vuelve a este flujo con el curso nuevo ya elegido
        const fecha = document.querySelector("[name=fecha]").value;
        const nota = document.querySelector("[name=nota]").value;
        const reabrir = cursoNuevoId => {
          modalInscribir({ alumnoId: alumnoSel ? alumnoSel.id : null, edicionId });
          const f = document.querySelector("[name=fecha]"); if (f && fecha) f.value = fecha;
          const n = document.querySelector("[name=nota]"); if (n) n.value = nota;
          const c = $("#ins-curso");
          if (c && cursoNuevoId) { c.value = cursoNuevoId; c.dispatchEvent(new Event("change")); }
        };
        modalCursoRapido(nuevo => reabrir(nuevo.id), () => reabrir(null));
        return;
      }
      if (!selCurso.value) { campoEd.classList.add("oculto"); return; }
      const eds = M.edicionesDeCurso(db, selCurso.value);
      campoEd.classList.remove("oculto");
      selEd.innerHTML = eds.map(e => `<option value="${e.id}">${esc(e.periodo || "sin fecha")}</option>`).join("") +
        `<option value="__nueva__">➕ Nueva edición…</option>`;
      if (!eds.length) { selEd.value = "__nueva__"; }
      selEd.dispatchEvent(new Event("change"));
    });
    selEd.addEventListener("change", () => {
      inputNuevaEd.classList.toggle("oculto", selEd.value !== "__nueva__");
      if (selEd.value === "__nueva__") inputNuevaEd.focus();
    });
  }

  // --- guardar la inscripción ---
  $("#form-inscribir").addEventListener("submit", e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    if (!alumnoSel) return toast("Elige o registra al alumno primero", "error");

    let edId = edicionId;
    if (!edId) {
      const selCurso = $("#ins-curso");
      const selEd = $("#ins-edicion");
      if (!selCurso.value || selCurso.value === "__nuevo__") return toast("Elige el curso", "error");
      if (selEd.value === "__nueva__") {
        const periodo = $("#ins-nueva-edicion").value.trim();
        if (!periodo) return toast("Escribe la fecha o periodo de la edición nueva", "error");
        edId = M.agregarEdicion(db, { cursoId: selCurso.value, periodo }).id;
      } else {
        edId = selEd.value;
      }
      if (!edId) return toast("Elige la edición", "error");
    }

    if (M.yaInscrito(db, alumnoSel.id, edId)) {
      return toast("⚠️ " + M.nombreCompleto(alumnoSel) + " ya está inscrito(a) en esa edición", "error");
    }

    const { promos } = M.inscribir(db, {
      alumnoId: alumnoSel.id,
      edicionId: edId,
      fecha: fd.get("fecha"),
      nota: fd.get("nota")
    });
    marcar();
    cerrarModal();
    if (promos.length) {
      toast(`✅ Inscrito · 🎁 Promoción aplicada: también se inscribió a ${promos.map(p => M.etiquetaEdicion(db, p.edicionId)).join(", ")}`, "promo");
    } else {
      toast("✅ Inscripción registrada");
    }
    render();
  });
}

// curso rápido (desde el flujo de inscripción)
function modalCursoRapido(alGuardar, alCancelar) {
  const caja = $("#modal-caja");
  caja.innerHTML = `
    <h3>➕ Curso nuevo (rápido)</h3>
    <form id="form-curso-rapido">
      <div class="form-grid">
        <div class="campo ancho-total"><label>Nombre del curso *</label><input name="nombre" required autofocus></div>
        <div class="campo"><label>Tipo</label>
          <select name="tipo"><option value="vivo">🔴 En vivo</option><option value="video">🎬 En video / antiguo</option></select>
        </div>
      </div>
      <div id="aviso-dup-rapido"></div>
      <div class="fila-botones">
        <button type="button" class="btn btn-suave" id="btn-cancelar-rapido">Volver</button>
        <button type="submit" class="btn btn-primario">Crear curso</button>
      </div>
    </form>
  `;
  $("#btn-cancelar-rapido").addEventListener("click", () => {
    cerrarModal();
    if (alCancelar) alCancelar();
  });
  let confirmado = false;
  $("#form-curso-rapido").addEventListener("submit", e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    const dup = M.cursoDuplicado(db, fd.get("nombre"));
    if (dup && !confirmado) {
      $("#aviso-dup-rapido").innerHTML = `<div class="aviso-duplicado">⚠️ Ya existe «${esc(dup.nombre)}». Vuelve a pulsar para crearlo de todos modos.</div>`;
      confirmado = true;
      return;
    }
    const c = M.agregarCurso(db, { nombre: fd.get("nombre"), tipo: fd.get("tipo") });
    marcar();
    toast("✅ Curso creado");
    cerrarModal();
    if (alGuardar) alGuardar(c);
  });
}
