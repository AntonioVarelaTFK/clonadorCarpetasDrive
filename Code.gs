const NOMBRE_CARPETA_INFORMES = "Informes de copia Drive";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("FormularioDestino")
    .setTitle("Copia desde carpeta compartida");
}

function iniciarCopia(datos) {
  try {
    const idOrigen = datos.idCarpeta.trim();
    let nombreDestino = datos.nombreDestino.trim() || "Copia desde compartido";
    let modo = datos.modo?.trim() || "auto";
    const carpetaOrigen = DriveApp.getFolderById(idOrigen);
    const nombreRaizOrigen = carpetaOrigen.getName();

    let carpetaDestino;
    let mensajeEstado = "";
    const carpetas = DriveApp.getFoldersByName(nombreDestino);

    let registro;

    if (carpetas.hasNext()) {
      carpetaDestino = carpetas.next();

      if (modo === "auto") {
        if (carpetaDestino.getFiles().hasNext() || carpetaDestino.getFolders().hasNext()) {
          const registroExistente = cargarRegistro(nombreDestino);
          if (Object.keys(registroExistente).length === 0) {
            modo = "continuar";
            registro = generarRegistroDesdeDestino(carpetaOrigen, carpetaDestino);
            mensajeEstado = `⚠ Carpeta "${nombreDestino}" tenía contenido pero sin registro previo. Se ha generado uno desde el contenido actual.`;
          } else {
            modo = "continuar";
          }
        } else {
          modo = "nuevo";
        }
      }

      if (modo === "nuevo" && (carpetaDestino.getFiles().hasNext() || carpetaDestino.getFolders().hasNext())) {
        modo = "continuar";
        mensajeEstado = `⚠ Carpeta "${nombreDestino}" ya tenía contenido. Cambiado a modo CONTINUAR.`;
      } else if (!mensajeEstado) {
        mensajeEstado = `Usando carpeta existente: "${nombreDestino}".`;
      }

    } else {
      carpetaDestino = DriveApp.createFolder(nombreDestino);
      mensajeEstado = `Carpeta creada: "${nombreDestino}".`;
    }

    registro = modo === "continuar" && typeof registro === "undefined" ? cargarRegistro(nombreDestino) : (registro || {});

    const nuevos = copiarRecursivoConProgreso(carpetaOrigen, carpetaDestino, registro, "", nombreRaizOrigen);

    let mensajeFinal = nuevos === 0
      ? "✅ Todo ya estaba copiado. No se han hecho cambios."
      : `🟡 Se han copiado ${nuevos} elementos nuevos.`;

    const nombreArchivoRegistro = generarNombreRegistro(nombreDestino);
    registro.__fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const archivoRegistro = guardarRegistro(registro, nombreArchivoRegistro);

    let advertencias = "";
    if (registro.__pesados?.length > 0) {
      advertencias += "⚠ Archivos grandes detectados:<br>" + registro.__pesados.map(p => "- " + p).join("<br>") + "<br><br>";
    }
    if (registro.__errores?.length > 0) {
      advertencias += "❌ Errores durante la copia:<br>" + registro.__errores.map(e => "- " + e).join("<br>") + "<br><br>";
    }

    let mensajeHTML = "";
    mensajeHTML += `<div>📁 ${nombreDestino}</div>`;
    mensajeHTML += `<div>${mensajeEstado}<br>${mensajeFinal}</div>`;

    if (advertencias.trim()) {
      mensajeHTML += `<div style="margin-top: 10px">${advertencias}</div>`;
    }

    return {
      mensaje: mensajeHTML,
      idDestino: carpetaDestino.getId(),
      nombreDestino: nombreDestino,
      urlRegistro: archivoRegistro.getUrl()
    };

  } catch (e) {
    Logger.log("Error general en iniciarCopia: " + e.message);
    return { mensaje: "❌ Ocurrió un error inesperado. Revisa los registros." };
  }
}


function copiarRecursivoConProgreso(origen, destino, registro, rutaBase, raiz) {
  const nombreActual = origen.getName();
  const rutaActual = rutaBase ? `${rutaBase}/${nombreActual}` : `/${nombreActual}`;

  if (!registro[rutaActual]) registro[rutaActual] = { archivos: {}, carpetas: {} };
  if (!registro.__errores) registro.__errores = [];
  if (!registro.__pesados) registro.__pesados = [];
  if (!registro.__raiz) registro.__raiz = raiz;

  let nuevosCopiados = 0;

  const archivosDestino = destino.getFiles();
  const mapaArchivosDestino = {};
  while (archivosDestino.hasNext()) {
    const archivo = archivosDestino.next();
    mapaArchivosDestino[archivo.getName()] = archivo;
  }

  const archivosOrigen = origen.getFiles();
  while (archivosOrigen.hasNext()) {
    const archivo = archivosOrigen.next();
    const nombre = archivo.getName();
    const size = archivo.getSize();
    const modificado = archivo.getLastUpdated().getTime();

    const yaRegistrado = registro[rutaActual].archivos[nombre];
    const mismoContenido = yaRegistrado &&
                           yaRegistrado.size === size &&
                           yaRegistrado.modificado === modificado;

    const existeEnDestino = !!mapaArchivosDestino[nombre]; 

    // 🔄 Si está registrado pero ha sido eliminado en destino, debe recopiarse
    const necesitaCopiar = !mismoContenido || !existeEnDestino;

    if (necesitaCopiar) {
      if (size > 100 * 1024 * 1024) {
        registro.__pesados.push(`${rutaActual}/${nombre} (${Math.round(size / 1024 / 1024)} MB)`);
      }
      try {
        if (existeEnDestino) {
          mapaArchivosDestino[nombre].setTrashed(true);
        }
        const copia = archivo.makeCopy(nombre, destino);
        registro[rutaActual].archivos[nombre] = {
          id: copia.getId(),
          size: size,
          modificado: modificado
        };
        nuevosCopiados++;
      } catch (e) {
        const msg = `Error al copiar archivo "${nombre}" en ${rutaActual}: ${e.message}`;
        Logger.log(msg);
        registro.__errores.push(msg);
      }
    }
  }

  const subcarpetasDestino = destino.getFolders();
  const mapaCarpetasDestino = {};
  while (subcarpetasDestino.hasNext()) {
    const c = subcarpetasDestino.next();
    mapaCarpetasDestino[c.getName()] = c.getId();
  }

  const subcarpetasOrigen = origen.getFolders();
  while (subcarpetasOrigen.hasNext()) {
    const sub = subcarpetasOrigen.next();
    const nombreSub = sub.getName();

    const idDestino = registro[rutaActual].carpetas[nombreSub];
    const existeSubEnDestino = idDestino && mapaCarpetasDestino[nombreSub]; 

    if (!idDestino || !existeSubEnDestino) { // 🔄 incluir creación si ha sido eliminada
      try {
        const nuevaSub = destino.createFolder(nombreSub);
        registro[rutaActual].carpetas[nombreSub] = nuevaSub.getId();
        nuevosCopiados++;
      } catch (e) {
        const msg = `Error al crear carpeta "${nombreSub}" en ${rutaActual}: ${e.message}`;
        Logger.log(msg);
        registro.__errores.push(msg);
        continue;
      }
    }

    try {
      const subDestino = DriveApp.getFolderById(registro[rutaActual].carpetas[nombreSub]);
      nuevosCopiados += copiarRecursivoConProgreso(sub, subDestino, registro, rutaActual, raiz);
    } catch (e) {
      const msg = `Error al continuar con subcarpeta "${nombreSub}" en ${rutaActual}: ${e.message}`;
      Logger.log(msg);
      registro.__errores.push(msg);
    }
  }

  return nuevosCopiados;
}

function verificarIntegridad(datos) {
  try {
    const idOrigen = datos.idCarpeta.trim();
    const nombreDestino = datos.nombreDestino.trim();
    const carpetaOrigen = DriveApp.getFolderById(idOrigen);

    const carpetas = DriveApp.getFoldersByName(nombreDestino);
    if (!carpetas.hasNext()) {
      return `❌ No se encontró la carpeta de destino "${nombreDestino}".`;
    }

    const carpetaDestino = carpetas.next();
    const registro = cargarRegistro(nombreDestino);
    const nombreRaizOrigen = carpetaOrigen.getName();
    const resultado = compararEstructuras(carpetaOrigen, carpetaDestino, "", registro, nombreRaizOrigen);
    //const resultado = compararEstructuras(carpetaOrigen, carpetaDestino, "", registro);
    

    if (resultado.trim() === "") {
      return "✅ Todo coincide: no se encontraron diferencias entre origen y destino.";
    } else {
      return `
        <strong>⚠ Verificación finalizada con incidencias:</strong><br><br>${resultado}
        <br><br>🔁 Puedes volver a lanzar la copia para intentar completar o actualizar los elementos.
        <br><br>🧹 Si deseas eliminar del destino los elementos que ya no están en el origen, usa el botón de <strong>Limpiar</strong>.`;
    }

  } catch (e) {
    return "❌ Error al verificar integridad: " + e.message;
  }
}

function compararEstructuras(origen, destino, rutaBase, registro, raiz) {
  const rutaActual = rutaBase ? `${rutaBase}/${origen.getName()}` : `/${raiz}`;
  let resultado = "";

  const registroRuta = registro[rutaActual]?.archivos || {};
  const carpetasRegistradas = registro[rutaActual]?.carpetas || {};

  const archivosOrigen = origen.getFiles();
  const mapaOrigen = {};
  const detallesOrigen = {};
  while (archivosOrigen.hasNext()) {
    const f = archivosOrigen.next();
    const nombre = f.getName();
    mapaOrigen[nombre] = f.getSize();
    detallesOrigen[nombre] = f.getLastUpdated().getTime();
  }

  const archivosDestino = destino.getFiles();
  const nombresDestino = new Set();
  while (archivosDestino.hasNext()) {
    nombresDestino.add(archivosDestino.next().getName());
  }

  for (const nombre in mapaOrigen) {
    const size = mapaOrigen[nombre];
    const modificado = detallesOrigen[nombre];

    if (!(nombre in registroRuta)) {
      resultado += `❌ Archivo no registrado aún: ${rutaActual}/${nombre}<br>`;
    } else {
      const reg = registroRuta[nombre];
      try {
        const archivoDestino = DriveApp.getFileById(reg.id);
        if (archivoDestino.isTrashed()) {
          resultado += `❌ Archivo eliminado (en papelera): ${rutaActual}/${nombre}<br>`;
        }
      } catch (e) {
        resultado += `❌ Archivo faltante en destino: ${rutaActual}/${nombre}<br>`;
        continue;
      }

      if (reg.size !== size) {
        resultado += `⚠ Tamaño cambiado: ${rutaActual}/${nombre} (antes: ${reg.size}, ahora: ${size})<br>`;
      } else if (reg.modificado !== modificado) {
        resultado += `⚠ Modificado desde la copia: ${rutaActual}/${nombre} (original: ${new Date(reg.modificado).toLocaleString()}, ahora: ${new Date(modificado).toLocaleString()})<br>`;
      }
    }
  }

  for (const nombre of nombresDestino) {
    if (!(nombre in mapaOrigen)) {
      resultado += `⚠ Archivo en destino que ya no está en origen: ${rutaActual}/${nombre}<br>`;
    }
  }

  // Verificar subcarpetas registradas eliminadas
  for (const nombre in carpetasRegistradas) {
    const id = carpetasRegistradas[nombre];
    try {
      const carpeta = DriveApp.getFolderById(id);
      if (carpeta.isTrashed()) {
        resultado += `❌ Subcarpeta eliminada (en papelera): ${rutaActual}/${nombre}<br>`;
      }
    } catch (e) {
      resultado += `❌ Subcarpeta faltante en destino: ${rutaActual}/${nombre}<br>`;
    }
  }

  const subcarpetasOrigen = origen.getFolders();
  const mapaSubOrigen = {};
  while (subcarpetasOrigen.hasNext()) {
    const sub = subcarpetasOrigen.next();
    mapaSubOrigen[sub.getName()] = sub;
  }

  const subcarpetasDestino = destino.getFolders();
  const mapaSubDestino = {};
  while (subcarpetasDestino.hasNext()) {
    const sub = subcarpetasDestino.next();
    mapaSubDestino[sub.getName()] = sub;
  }

  for (const nombre in mapaSubOrigen) {
    if (!(nombre in mapaSubDestino)) {
      resultado += `❌ Falta subcarpeta en destino: ${rutaActual}/${nombre}<br>`;
    } else {
      resultado += compararEstructuras(mapaSubOrigen[nombre], mapaSubDestino[nombre], rutaActual, registro, raiz);
    }
  }

  for (const nombre in mapaSubDestino) {
    if (!(nombre in mapaSubOrigen)) {
      resultado += `⚠ Carpeta en destino que ya no está en origen: ${rutaActual}/${nombre}<br>`;
    }
  }

  return resultado;
}


function generarNombreRegistro(nombreDestino) {
  const ahora = new Date();
  const pad = n => n.toString().padStart(2, '0');
  const marca = `${ahora.getFullYear()}-${pad(ahora.getMonth() + 1)}-${pad(ahora.getDate())}_${pad(ahora.getHours())}-${pad(ahora.getMinutes())}`;
  const nombreLimpiado = nombreDestino.replace(/[^\w\d-_]/g, "_");
  return `progreso_copia_drive_${nombreLimpiado}_${marca}.json`;
}

function guardarRegistro(registro, nombreArchivo) {
  try {
    const json = JSON.stringify(registro, null, 2);
    const carpetaInformes = obtenerOCrearCarpeta(NOMBRE_CARPETA_INFORMES);
    const archivo = carpetaInformes.createFile(nombreArchivo, json, MimeType.PLAIN_TEXT);
    return archivo;
  } catch (e) {
    Logger.log("Error al guardar el registro: " + e.message);
    return null;
  }
}

function cargarRegistro(nombreDestino) {
  try {
    const carpetaInformes = obtenerOCrearCarpeta(NOMBRE_CARPETA_INFORMES);
    const archivos = carpetaInformes.getFiles();
    let ultimoArchivo = null;
    let ultimaFecha = 0;

    const nombreLimpiado = nombreDestino.replace(/[^\w\d-_]/g, "_");
    const prefijo = `progreso_copia_drive_${nombreLimpiado}`;

    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombre = archivo.getName();

      if (!nombre.startsWith(prefijo)) continue;

      const fecha = archivo.getLastUpdated().getTime();
      if (fecha > ultimaFecha) {
        ultimaFecha = fecha;
        ultimoArchivo = archivo;
      }
    }

    if (!ultimoArchivo) return {};
    const contenido = ultimoArchivo.getBlob().getDataAsString();
    return JSON.parse(contenido);
  } catch (e) {
      Logger.log("Error al cargar el registro: " + e.message);
      return {};
  }
}

function obtenerOCrearCarpeta(nombre) {
  const carpetas = DriveApp.getFoldersByName(nombre);
  return carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(nombre);
}

function limpiarDestino(datos) {
  try {
    const idOrigen = datos.idCarpeta.trim();
    const nombreDestino = datos.nombreDestino.trim();
    const carpetaOrigen = DriveApp.getFolderById(idOrigen);

    const carpetas = DriveApp.getFoldersByName(nombreDestino);
    if (!carpetas.hasNext()) {
      return `❌ No se encontró la carpeta de destino "${nombreDestino}".`;
    }

    const carpetaDestino = carpetas.next();
    const registroEliminados = [];

    eliminarSobrantes(carpetaOrigen, carpetaDestino, "", registroEliminados);

    if (registroEliminados.length === 0) {
      return "✅ No había elementos sobrantes. Nada se ha eliminado.";
    } else {
      const resumen = registroEliminados.map(e => "- " + e).join("<br>");
      return `🧹 Se han eliminado ${registroEliminados.length} elementos sobrantes:<br><br>${resumen}`;
    }

  } catch (e) {
    return "❌ Error durante la limpieza: " + e.message;
  }
}


function eliminarSobrantes(origen, destino, rutaBase, eliminados) {
  const rutaActual = rutaBase + "/" + origen.getName();

  // Archivos
  const archivosOrigen = origen.getFiles();
  const setOrigen = new Set();
  while (archivosOrigen.hasNext()) {
    setOrigen.add(archivosOrigen.next().getName());
  }

  const archivosDestino = destino.getFiles();
  while (archivosDestino.hasNext()) {
    const f = archivosDestino.next();
    const nombre = f.getName();
    if (!setOrigen.has(nombre)) {
      try {
        f.setTrashed(true);
        eliminados.push(`Archivo eliminado del destino: ${rutaActual}/${nombre}`);
      } catch (e) {
        eliminados.push(`❌ Error al eliminar archivo: ${rutaActual}/${nombre} → ${e.message}`);
      }
    }
  }

  // Subcarpetas
  const subcarpetasOrigen = origen.getFolders();
  const setSubOrigen = new Set();
  while (subcarpetasOrigen.hasNext()) {
    setSubOrigen.add(subcarpetasOrigen.next().getName());
  }

  const subcarpetasDestino = destino.getFolders();
  while (subcarpetasDestino.hasNext()) {
    const sub = subcarpetasDestino.next();
    const nombre = sub.getName();
    if (!setSubOrigen.has(nombre)) {
      try {
        sub.setTrashed(true);
        eliminados.push(`Carpeta eliminada del destino: ${rutaActual}/${nombre}`);
      } catch (e) {
        eliminados.push(`❌ Error al eliminar carpeta: ${rutaActual}/${nombre} → ${e.message}`);
      }
    } else {
      const subOrigen = origen.getFoldersByName(nombre);
      if (subOrigen.hasNext()) {
        eliminarSobrantes(subOrigen.next(), sub, rutaActual, eliminados);
      }
    }
  }
}

function generarRegistroDesdeDestino(origen, destino) {
  const registro = { __errores: [], __pesados: [], __raiz: origen.getName() };
  construirRegistroDesdeDestino(origen, destino, registro, "");
  return registro;
}

function construirRegistroDesdeDestino(origen, destino, registro, rutaBase) {
  const nombreActual = origen.getName();
  const rutaActual = rutaBase ? `${rutaBase}/${nombreActual}` : `/${nombreActual}`;
  registro[rutaActual] = { archivos: {}, carpetas: {} };

  const archivosOrigen = origen.getFiles();
  const mapaOrigen = {};
  while (archivosOrigen.hasNext()) {
    const a = archivosOrigen.next();
    mapaOrigen[a.getName()] = a.getSize();
  }

  const archivosDestino = destino.getFiles();
  while (archivosDestino.hasNext()) {
    const f = archivosDestino.next();
    const nombre = f.getName();
    const size = f.getSize();
    if (nombre in mapaOrigen && mapaOrigen[nombre] === size) {
      registro[rutaActual].archivos[nombre] = {
        id: f.getId(),
        size: size,
        modificado: null  // Se ignorará en la siguiente comprobación
      };
    }
  }

  const subcarpetasOrigen = origen.getFolders();
  const mapaSubOrigen = {};
  while (subcarpetasOrigen.hasNext()) {
    const sub = subcarpetasOrigen.next();
    mapaSubOrigen[sub.getName()] = sub;
  }

  const subcarpetasDestino = destino.getFolders();
  while (subcarpetasDestino.hasNext()) {
    const sub = subcarpetasDestino.next();
    const nombre = sub.getName();
    if (nombre in mapaSubOrigen) {
      registro[rutaActual].carpetas[nombre] = sub.getId();
      construirRegistroDesdeDestino(mapaSubOrigen[nombre], sub, registro, rutaActual);
    }
  }
}

/**
 * Comprime una carpeta destino previamente verificada en bloques de hasta 100 MB.
 * Los archivos ZIP se almacenan en una carpeta /Backups ZIP dentro del mismo Drive.
 * Se genera un log en la carpeta de informes con los archivos incluidos en cada ZIP.
 */
function comprimirCarpetaDestino(nombreDestino) {
  const carpetaDestino = DriveApp.getFoldersByName(nombreDestino);
  if (!carpetaDestino.hasNext()) throw new Error(`No se encuentra la carpeta destino: ${nombreDestino}`);
  const carpeta = carpetaDestino.next();

  const parent = carpeta.getParents().hasNext() ? carpeta.getParents().next() : DriveApp.getRootFolder();
  const backupsZipFolder = parent.getFoldersByName("Backups ZIP").hasNext()
    ? parent.getFoldersByName("Backups ZIP").next()
    : parent.createFolder("Backups ZIP");

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const baseNombreZip = `backup_${nombreDestino.replace(/[^\w\d-_]/g, '_')}_${fecha}`;

  const archivos = obtenerTodosLosArchivosConRuta(carpeta, "");
  const BLOQUE_MB = 100;
  const BLOQUE_BYTES = BLOQUE_MB * 1024 * 1024;

  // Paso previo: estimar número de bloques
  let total = 0;
  let bloquesEstimados = 0;
  for (const obj of archivos) {
    const size = obj.file.getSize();
    total += size;
    if (total > BLOQUE_BYTES) {
      bloquesEstimados++;
      total = size;
    }
  }
  if (total > 0) bloquesEstimados++; // último bloque

  // Ahora hacer compresión real
  let bloque = [];
  let numZip = 1;
  let zipsCreados = [];
  let resumenLog = [];
  let advertencias = [];
  total = 0;

  for (const obj of archivos) {
    const size = obj.file.getSize();
    if (total + size > BLOQUE_BYTES && bloque.length > 0) {
      const nombreZip = `${baseNombreZip}_part${numZip}_of_${bloquesEstimados}.zip`;
      const zip = crearZipDesdeArchivos(bloque, nombreZip, backupsZipFolder);
      zipsCreados.push(zip.getUrl());
      resumenLog.push(`ZIP ${numZip} (${bloque.length} archivos)`);
      bloque = [];
      total = 0;
      numZip++;
    }
    bloque.push(obj);
    total += size;
  }

  if (bloque.length > 0) {
    const nombreZip = `${baseNombreZip}_part${numZip}_of_${bloquesEstimados}.zip`;
    const zip = crearZipDesdeArchivos(bloque, nombreZip, backupsZipFolder, advertencias);
    zipsCreados.push(zip.getUrl());
    resumenLog.push(`ZIP ${numZip} (${bloque.length} archivos)`);
  }

  const log = {
    destino: nombreDestino,
    fecha: fecha,
    zips: zipsCreados,
    resumen: resumenLog,
    advertencias: advertencias
  };

  const nombreLog = `log_zip_${baseNombreZip}.json`;
  const logFile = guardarRegistro(log, nombreLog);

  return {
    mensaje: `✅ Se han creado ${zipsCreados.length} archivo(s) ZIP con estructura.`,
    zips: zipsCreados,
    urlLog: logFile.getUrl(),
    carpetaZipsUrl: backupsZipFolder.getUrl()
  };
}


/**
 * Recorre recursivamente una carpeta y devuelve todos los archivos
 */
function obtenerTodosLosArchivosConRuta(folder, rutaBase) {
  let archivos = [];
  const nombreActual = folder.getName();
  const rutaActual = rutaBase ? `${rutaBase}/${nombreActual}` : nombreActual;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    archivos.push({ file: f, ruta: `${rutaActual}/${f.getName()}` });
  }

  const subcarpetas = folder.getFolders();
  while (subcarpetas.hasNext()) {
    const sub = subcarpetas.next();
    archivos = archivos.concat(obtenerTodosLosArchivosConRuta(sub, rutaActual));
  }

  return archivos;
}

/**
 * Recibe una lista de archivos y los comprime usando Utilities.zip
 */
function crearZipDesdeArchivos(archivos, nombreZip, destinoFolder, advertencias) {
  const blobs = archivos.map(obj => {
    const file = obj.file;
    const ruta = obj.ruta;
    const mimeType = file.getMimeType();
    const fileId = file.getId();

    try {
      if (mimeType === MimeType.GOOGLE_DOCS || mimeType === MimeType.GOOGLE_SHEETS || mimeType === MimeType.GOOGLE_SLIDES) {
        const exportMime =
          mimeType === MimeType.GOOGLE_DOCS ? MimeType.MICROSOFT_WORD :
          mimeType === MimeType.GOOGLE_SHEETS ? MimeType.MICROSOFT_EXCEL :
          MimeType.MICROSOFT_POWERPOINT;

        const extension =
          mimeType === MimeType.GOOGLE_DOCS ? ".docx" :
          mimeType === MimeType.GOOGLE_SHEETS ? ".xlsx" :
          ".pptx";

        const url = `https://www.googleapis.com/drive/v2/files/${fileId}/export?mimeType=${encodeURIComponent(exportMime)}`;
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: `Bearer ${token}` },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
          throw new Error(`Código ${response.getResponseCode()}: ${response.getContentText()}`);
        }

        return Utilities.newBlob(response.getBlob().getBytes(), exportMime, ruta.replace(/\.\w+$/, extension));

      } else if (mimeType === MimeType.GOOGLE_FORMS) {
        const url = file.getUrl();
        const acceso = `[InternetShortcut]\nURL=${url}`;
        const blob = Utilities.newBlob(acceso, MimeType.PLAIN_TEXT, ruta.replace(/\.gform$/, ".url"));
        advertencias.push(`⚠ Formulario incluido como acceso directo: ${ruta}`);
        return blob;

      } else {
        return file.getBlob().setName(ruta);
      }

    } catch (e) {
      Logger.log(`⚠ Error al exportar archivo ${file.getName()}: ${e.message}`);
      advertencias.push(`⚠ Error al exportar archivo ${ruta}. Error: ${e.message}. Mimetype: ${mimeType}. Añadido como blob original.`);
      return file.getBlob().setName(ruta);
    }
  });

  const zipBlob = Utilities.zip(blobs, nombreZip);
  return destinoFolder.createFile(zipBlob);
}


/**
 * Calcula el número aproximado de ficheros zip a generar
 */
function calcularEstimacionCompresion(nombreDestino) {
  const carpetaDestino = DriveApp.getFoldersByName(nombreDestino);
  if (!carpetaDestino.hasNext()) throw new Error(`No se encuentra la carpeta destino: ${nombreDestino}`);
  const carpeta = carpetaDestino.next();

  const archivos = obtenerTodosLosArchivosConRuta(carpeta, "");
  const BLOQUE_MB = 100;
  const BLOQUE_BYTES = BLOQUE_MB * 1024 * 1024;

  let totalBytes = 0;
  for (const obj of archivos) {
    totalBytes += obj.file.getSize();
  }

  const numZips = Math.ceil(totalBytes / BLOQUE_BYTES);
  const gbTotales = (totalBytes / (1024 ** 3)).toFixed(2);

  return {
    mensaje: `📦 Estimado: ${gbTotales} GB en ${numZips} archivo(s) ZIP de hasta ${BLOQUE_MB} MB.`,
    totalGB: gbTotales,
    partes: numZips
  };
}

