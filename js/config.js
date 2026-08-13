// ============================================================
// CONFIGURACIÓN DE LA APP
//
// Modo en uso: carpeta sincronizada de Google Drive.
// La app abre la carpeta que elijas y trabaja sobre el archivo
// "localFileName" que está adentro. No requiere ningún registro
// ni cuenta de desarrollador.
//
// El modo OneDrive-por-internet quedó construido pero apagado:
// Microsoft ya no permite registrar apps con cuentas personales
// sin un "directorio" de Azure. Si algún día se crea uno, basta
// poner aquí el clientId y el botón reaparece solo.
// ============================================================
window.APP_CONFIG = {
  // ID de aplicación (cliente) del registro en Microsoft Entra.
  // Vacío = modo OneDrive apagado.
  clientId: "",

  // No cambiar: permite cuentas personales de Microsoft y de trabajo
  authority: "https://login.microsoftonline.com/common",

  // Ruta del archivo de datos dentro de OneDrive (desde la raíz)
  dataFilePath: "PRIOSAN/BaseAlumnos/datos-alumnos.json",

  // Carpeta donde se dejan respaldos con fecha en cada guardado
  backupFolderPath: "PRIOSAN/BaseAlumnos/respaldos",

  // Nombre del archivo de datos dentro de la carpeta de Google Drive
  localFileName: "datos-alumnos.json",

  // Cuántos respaldos con fecha conservar en la subcarpeta "respaldos".
  // Al guardar, los más antiguos que sobrepasen este número se borran solos.
  // Pon 0 para desactivar la limpieza y conservarlos todos.
  maxRespaldos: 30
};
