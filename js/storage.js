// ============================================================
// CAPA DE GUARDADO
// Tres proveedores intercambiables:
//  - GraphStorage: OneDrive vía internet (Microsoft Graph + MSAL)
//  - FolderStorage: carpeta local (File System Access API)
//  - FileStorage: cargar / descargar manual (respaldo de emergencia)
// Todos exponen: nombre, cargar() -> objeto|null, guardar(db) -> void
// ============================================================

const CFG = window.APP_CONFIG;

// Permisos que la app pide a Microsoft. "Files.ReadWrite.All" es necesario
// (en vez del más acotado "Files.ReadWrite") porque la segunda persona abre
// el archivo desde el OneDrive de otra cuenta, vía "compartidos conmigo".
const PERMISOS = ["Files.ReadWrite.All", "User.Read"];

function marcaDeTiempo() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}

// ---------------- OneDrive (Microsoft Graph) ----------------
export class GraphStorage {
  constructor() {
    this.nombre = "OneDrive (internet)";
    this.tipo = "graph";
    this.msal = null;
    this.cuenta = null;
    this.eTag = null;      // para detectar si alguien más guardó desde que cargamos
    this.driveId = null;   // en qué OneDrive vive el archivo (propio o compartido)
    this.itemId = null;    // identificador del archivo
    this.parentId = null;  // carpeta que lo contiene (para los respaldos)
    this.esCompartido = false;
  }

  static configurado() {
    return !!(CFG.clientId && CFG.clientId.trim());
  }

  // Dirección de retorno que Microsoft exige registrar. Se quita "index.html"
  // para que sea idéntica tanto si se abre ".../alumnos/" como ".../alumnos/index.html".
  static redirectUri() {
    const ruta = window.location.pathname.replace(/index\.html$/, "");
    return window.location.origin + ruta;
  }

  async conectar() {
    if (!window.msal) throw new Error("No se pudo cargar la librería de Microsoft (¿sin internet?).");
    this.msal = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: CFG.authority,
        redirectUri: GraphStorage.redirectUri()
      },
      cache: { cacheLocation: "localStorage" }
    });
    await this.msal.handleRedirectPromise();
    const cuentas = this.msal.getAllAccounts();
    if (cuentas.length) {
      this.cuenta = cuentas[0];
    } else {
      const res = await this.msal.loginPopup({ scopes: PERMISOS });
      this.cuenta = res.account;
    }
    return this.cuenta.username;
  }

  async _token() {
    const req = { scopes: PERMISOS, account: this.cuenta };
    try {
      const res = await this.msal.acquireTokenSilent(req);
      return res.accessToken;
    } catch {
      const res = await this.msal.acquireTokenPopup(req);
      return res.accessToken;
    }
  }

  async _fetch(ruta, opciones = {}) {
    const token = await this._token();
    opciones.headers = Object.assign({ Authorization: "Bearer " + token }, opciones.headers || {});
    return fetch("https://graph.microsoft.com/v1.0" + ruta, opciones);
  }

  // Guarda las señas del archivo encontrado (en qué OneDrive vive y cuál es).
  _fijar(driveId, itemId, parentId, eTag, compartido) {
    this.driveId = driveId;
    this.itemId = itemId;
    this.parentId = parentId;
    this.eTag = eTag || null;
    this.esCompartido = compartido;
    this.nombre = compartido ? "OneDrive (compartido)" : "OneDrive (internet)";
  }

  // Localiza el archivo de datos: primero en el OneDrive propio y,
  // si no está, entre los elementos que otra cuenta compartió contigo.
  async _resolver() {
    if (this.itemId) return;
    const nombreArchivo = CFG.dataFilePath.split("/").pop();
    const nombreCarpeta = CFG.dataFilePath.split("/").slice(-2)[0];

    // 1) En el OneDrive propio
    const propio = await this._fetch(encodeURI("/me/drive/root:/" + CFG.dataFilePath));
    if (propio.ok) {
      const info = await propio.json();
      this._fijar(info.parentReference.driveId, info.id, info.parentReference.id, info.eTag, false);
      return;
    }
    if (propio.status !== 404) throw new Error("Error al leer OneDrive (" + propio.status + ")");

    // 2) Entre lo compartido conmigo
    const comp = await this._fetch("/me/drive/sharedWithMe");
    if (comp.ok) {
      const lista = (await comp.json()).value || [];

      // 2a) El archivo compartido directamente
      const arch = lista.find(x => x.name === nombreArchivo && x.remoteItem);
      if (arch) {
        const r = arch.remoteItem;
        this._fijar(r.parentReference.driveId, r.id, r.parentReference.id, r.eTag, true);
        return;
      }

      // 2b) La carpeta compartida que lo contiene
      const carpeta = lista.find(x => x.name === nombreCarpeta && x.remoteItem && x.remoteItem.folder);
      if (carpeta) {
        const dId = carpeta.remoteItem.parentReference.driveId;
        const fId = carpeta.remoteItem.id;
        const hijo = await this._fetch(`/drives/${dId}/items/${fId}:/${encodeURIComponent(nombreArchivo)}`);
        if (hijo.ok) {
          const info = await hijo.json();
          this._fijar(dId, info.id, fId, info.eTag, true);
          return;
        }
      }
    }
    this.itemId = null; // no existe en ningún lado todavía
  }

  async cargar() {
    await this._resolver();
    if (!this.itemId) return null;
    const cont = await this._fetch(`/drives/${this.driveId}/items/${this.itemId}/content`);
    if (!cont.ok) throw new Error("Error al descargar el archivo (" + cont.status + ")");
    return await cont.json();
  }

  async guardar(db) {
    const cuerpo = JSON.stringify(db, null, 2);
    let res;

    if (this.itemId) {
      // Protección contra pisarse los cambios: solo guarda si nadie más
      // ha modificado el archivo desde que lo cargamos.
      const headers = { "Content-Type": "application/json" };
      if (this.eTag) headers["If-Match"] = this.eTag;
      res = await this._fetch(`/drives/${this.driveId}/items/${this.itemId}/content`,
        { method: "PUT", headers, body: cuerpo });
    } else {
      // Primera vez: se crea en el OneDrive propio
      res = await this._fetch(encodeURI("/me/drive/root:/" + CFG.dataFilePath + ":/content"),
        { method: "PUT", headers: { "Content-Type": "application/json" }, body: cuerpo });
    }

    if (res.status === 412) {
      throw new Error("CONFLICTO: alguien más guardó cambios después de que tú abriste la base. " +
        "Para no perder su trabajo, exporta un respaldo (Exportar → Respaldo JSON), recarga la página y vuelve a capturar tus cambios.");
    }
    if (!res.ok) throw new Error("Error al guardar en OneDrive (" + res.status + ")");

    const info = await res.json();
    this._fijar(info.parentReference.driveId, info.id, info.parentReference.id, info.eTag, this.esCompartido);

    // Respaldo con fecha junto al archivo (si falla, no interrumpe el guardado)
    try {
      await this._fetch(
        `/drives/${this.driveId}/items/${this.parentId}:/respaldos/datos-alumnos-${marcaDeTiempo()}.json:/content`,
        { method: "PUT", headers: { "Content-Type": "application/json" }, body: cuerpo });
    } catch { /* respaldo opcional */ }
  }
}

// ---------------- Carpeta local (File System Access) ----------------
export class FolderStorage {
  constructor() {
    this.nombre = "Carpeta sincronizada (Google Drive)";
    this.tipo = "carpeta";
    this.dir = null;
    this.revBase = null; // versión que traía el archivo cuando lo cargamos
  }

  static disponible() {
    return "showDirectoryPicker" in window;
  }

  async conectar() {
    this.dir = await window.showDirectoryPicker({ mode: "readwrite" });
    return this.dir.name;
  }

  async cargar() {
    try {
      const fh = await this.dir.getFileHandle(CFG.localFileName);
      const f = await fh.getFile();
      const datos = JSON.parse(await f.text());
      this.revBase = datos?.meta?.rev ?? 0;
      return datos;
    } catch (e) {
      if (e.name === "NotFoundError") return null;
      throw e;
    }
  }

  // Lee la versión que hay ahora mismo en el disco, sin cargar toda la base.
  async _revEnDisco() {
    try {
      const fh = await this.dir.getFileHandle(CFG.localFileName);
      const f = await fh.getFile();
      return JSON.parse(await f.text())?.meta?.rev ?? 0;
    } catch (e) {
      if (e.name === "NotFoundError") return null;
      throw e;
    }
  }

  // Conserva solo los respaldos más recientes y borra los que sobran.
  //
  // Reglas de seguridad, a propósito estrictas:
  //  - Solo borra archivos cuyo nombre calce EXACTAMENTE con el patrón que
  //    genera la app ("datos-alumnos-AAAAMMDD-HHMM.json"). Cualquier otro
  //    archivo que haya en la carpeta se queda intacto.
  //  - Nunca toca subcarpetas.
  //  - Nunca toca el archivo vivo (que vive fuera de "respaldos").
  //  - Los nombres llevan fecha de ancho fijo, así que ordenarlos
  //    alfabéticamente equivale a ordenarlos por fecha.
  async _limpiarRespaldos(dirR) {
    const max = CFG.maxRespaldos;
    if (!max || max < 1) return 0; // 0 = conservar todos
    const patron = /^datos-alumnos-\d{8}-\d{4}\.json$/;
    const nombres = [];
    for await (const [nombre, handle] of dirR.entries()) {
      if (handle.kind === "file" && patron.test(nombre)) nombres.push(nombre);
    }
    if (nombres.length <= max) return 0;
    nombres.sort().reverse();              // más reciente primero
    const sobrantes = nombres.slice(max);  // los más viejos
    let borrados = 0;
    for (const n of sobrantes) {
      try { await dirR.removeEntry(n); borrados++; } catch { /* si uno falla, seguimos */ }
    }
    return borrados;
  }

  // Guarda una exportación dentro de la subcarpeta "exportaciones" de la
  // misma carpeta de Google Drive. Devuelve la ruta para avisar al usuario.
  async guardarExport(nombre, contenido) {
    const dirExp = await this.dir.getDirectoryHandle("exportaciones", { create: true });
    const fh = await dirExp.getFileHandle(nombre, { create: true });
    const w = await fh.createWritable();
    await w.write(contenido);
    await w.close();
    return `${this.dir.name}/exportaciones/${nombre}`;
  }

  async guardar(db) {
    const cuerpo = JSON.stringify(db, null, 2);

    // Protección contra pisarse los cambios: si OneDrive ya sincronizó un
    // guardado de la otra persona, el archivo en disco trae una versión más
    // nueva que la que cargamos, y no debemos sobreescribirla.
    if (this.revBase !== null) {
      const enDisco = await this._revEnDisco();
      if (enDisco !== null && enDisco !== this.revBase) {
        throw new Error("CONFLICTO: el archivo cambió desde que lo abriste (seguramente la otra persona guardó y OneDrive ya sincronizó). " +
          "Para no perder su trabajo, exporta un respaldo (Exportar → Respaldo JSON), vuelve a abrir la base y captura de nuevo tus cambios.");
      }
    }

    const fh = await this.dir.getFileHandle(CFG.localFileName, { create: true });
    const w = await fh.createWritable();
    await w.write(cuerpo);
    await w.close();
    this.revBase = db?.meta?.rev ?? 0;

    // Respaldo con fecha en subcarpeta "respaldos"
    this.ultimaLimpieza = 0;
    try {
      const dirR = await this.dir.getDirectoryHandle("respaldos", { create: true });
      const fhR = await dirR.getFileHandle("datos-alumnos-" + marcaDeTiempo() + ".json", { create: true });
      const wR = await fhR.createWritable();
      await wR.write(cuerpo);
      await wR.close();
      this.ultimaLimpieza = await this._limpiarRespaldos(dirR);
    } catch { /* respaldo opcional: nunca debe impedir el guardado */ }
  }
}

// ---------------- Archivo manual (cargar / descargar) ----------------
export class FileStorage {
  constructor() {
    this.nombre = "Archivo manual";
    this.tipo = "archivo";
  }

  async cargar() {
    return new Promise((resolve, reject) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".json,application/json";
      input.onchange = async () => {
        try {
          const f = input.files[0];
          if (!f) return resolve(null);
          resolve(JSON.parse(await f.text()));
        } catch (e) { reject(e); }
      };
      input.click();
    });
  }

  async guardar(db) {
    const blob = new Blob([JSON.stringify(db, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "datos-alumnos-" + marcaDeTiempo() + ".json";
    a.click();
    URL.revokeObjectURL(a.href);
  }
}
