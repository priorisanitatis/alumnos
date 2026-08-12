// ============================================================
// CONFIGURACIÓN DE LA APP
// Rellena "clientId" con el ID de aplicación que te dé el
// registro en Microsoft (portal.azure.com → App registrations).
// Mientras esté vacío, el botón "Conectar con OneDrive" mostrará
// las instrucciones en lugar de conectar.
// ============================================================
window.APP_CONFIG = {
  // ID de aplicación (cliente) del registro en Microsoft Entra
  clientId: "",

  // No cambiar: permite cuentas personales de Microsoft y de trabajo
  authority: "https://login.microsoftonline.com/common",

  // Ruta del archivo de datos dentro de OneDrive (desde la raíz)
  dataFilePath: "PRIOSAN/BaseAlumnos/datos-alumnos.json",

  // Carpeta donde se dejan respaldos con fecha en cada guardado
  backupFolderPath: "PRIOSAN/BaseAlumnos/respaldos",

  // Nombre del archivo de datos en el modo "carpeta local"
  localFileName: "datos-alumnos.json"
};
