const NOMBRE_CARPETA_INFORMES = "Informes de copia Drive";
const LOTE_TAMAÑO = 7; // Procesar 7 tareas por ejecución del trigger (puedes ajustar esto)
const TRIGGER_DELAY_MS = 50 * 1000; // 50 segundos
const MAX_INTENTOS_POR_TAREA = 3;

// Claves para PropertiesService
const PS_KEYS = {
  BATCH_ACTIVE: 'BATCH_PROCESS_ACTIVE',
  LOG_FILENAME: 'BATCH_LOG_FILENAME',
  TASK_INDEX: 'BATCH_CURRENT_TASK_INDEX',
  TRIGGER_ID: 'BATCH_CURRENT_TRIGGER_ID',
  DESTINATION_FOLDER_ID: 'BATCH_DESTINATION_FOLDER_ID', // ID de la carpeta destino raíz
  SOURCE_FOLDER_ID: 'BATCH_SOURCE_FOLDER_ID' // ID de la carpeta origen raíz
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("FormularioDestino")
    .setTitle("Copia desde carpeta compartida por Lotes");
}

// ==========================================================================
// FUNCION PRINCIPAL INVOCADA DESDE HTML PARA INICIAR EL PROCESO
// ==========================================================================
function iniciarCopiaPorLotes(datos) {
  try {
    // Limpiar cualquier proceso de lote anterior para esta sesión de usuario
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered'); // Limpia triggers antiguos

    const idOrigen = datos.idCarpeta.trim();
    let nombreDestino = datos.nombreDestino.trim() || "Copia por lotes desde compartido";
    // El modo "forzar nuevo" o "continuar" se manejará dentro de la lógica de planificación.

    const carpetaOrigen = DriveApp.getFolderById(idOrigen);
    const nombreRaizOrigen = carpetaOrigen.getName();

    let carpetaDestino;
    let mensajeEstadoGlobal = "";
    const carpetasDestinoExistentes = DriveApp.getFoldersByName(nombreDestino);

    if (carpetasDestinoExistentes.hasNext()) {
      carpetaDestino = carpetasDestinoExistentes.next();
      // Si hay más con el mismo nombre, usamos la primera y advertimos (o podrías tener una lógica más compleja)
      if (carpetasDestinoExistentes.hasNext()) {
        Logger.log(`ADVERTENCIA: Múltiples carpetas destino encontradas con el nombre "${nombreDestino}". Usando la primera: ${carpetaDestino.getId()}`);
      }
      mensajeEstadoGlobal = `Usando carpeta destino existente: "${nombreDestino}".`;
    } else {
      carpetaDestino = DriveApp.createFolder(nombreDestino);
      mensajeEstadoGlobal = `Carpeta destino creada: "${nombreDestino}".`;
    }
    
    const idCarpetaDestino = carpetaDestino.getId();
    PropertiesService.getUserProperties().setProperty(PS_KEYS.DESTINATION_FOLDER_ID, idCarpetaDestino);
    PropertiesService.getUserProperties().setProperty(PS_KEYS.SOURCE_FOLDER_ID, idOrigen);

    // Cargar el registro más reciente si existe para esta carpeta destino
    let dataRegistro = cargarRegistro(nombreDestino);
    let registro =  dataRegistro.registroJson;
    ////////////////////
    Logger.log(`Registro cargado para "${nombreDestino}". ¿Tiene __raiz?: ${registro.hasOwnProperty('__raiz')}. Contenido de __progreso.mensajeActual: ${registro.__progreso ? registro.__progreso.mensajeActual : 'N/A'}`);

    // Determinar si el registro cargado es un placeholder de "no encontrado" o "error al cargar"
    const esRegistroPlaceholder = registro.__progreso && registro.__progreso.mensajeActual &&
                                  (registro.__progreso.mensajeActual.includes("No se encontró registro previo.") ||
                                   registro.__progreso.mensajeActual.includes("Error al cargar registro."));

    if (Object.keys(registro).length === 0 || esRegistroPlaceholder) {
      // Si está completamente vacío o es el placeholder de error/no encontrado de cargarRegistro
      Logger.log(`Registro para "${nombreDestino}" está vacío o es un placeholder. Inicializando nuevo registro.`);
      registro = {
          // __raiz se establece más abajo de forma garantizada
          __fecha: null, // Se establecerá más abajo
          __estado_proceso: "planificando",
          __progreso: { totalTareas: 0, tareasCompletadas: 0, erroresEnLote: 0, mensajeActual: "Iniciando planificación..." },
          __pendientes: [],
          __errores: [],
          __pesados: []
      };
    } else {
        Logger.log(`Usando registro existente para "${nombreDestino}". Limpiando para replanificación.`);
        // Limpiar pendientes y progreso para una nueva planificación, pero mantener la estructura de carpetas/archivos ya conocida.
        registro.__pendientes = []; // Limpiar siempre para una nueva planificación desde UI
        registro.__progreso = { 
            totalTareas: 0, 
            tareasCompletadas: 0, // Reseteamos para la nueva planificación
            erroresEnLote: 0, 
            mensajeActual: "Re-planificando tareas desde registro existente..." 
        };
    }

    // Asegurar SIEMPRE que __raiz y __fecha estén actualizados para la ejecución actual.
    // Esto también maneja el caso donde un registro antiguo pudiera no tener __raiz.
    registro.__raiz = nombreRaizOrigen;
    registro.__fecha = new Date().toISOString();
    registro.__estado_proceso = "planificando"; // Siempre estamos planificando aquí
    Logger.log(`__raiz para "${nombreDestino}" establecida/actualizada a: "${registro.__raiz}"`);

    // --- Fase de Planificación ---
    let listaTareas = [];
    // generarListaDeTareas usará el 'registro' con __raiz ya asegurado.
    // También usará 'registro' para ver la estructura de destino ya conocida si existe.
    generarListaDeTareas(carpetaOrigen, carpetaDestino, nombreRaizOrigen, registro, listaTareas);
    
    registro.__pendientes = listaTareas;
    registro.__progreso.totalTareas = listaTareas.length;
    registro.__progreso.mensajeActual = `Planificación completa. ${listaTareas.length} tareas pendientes.`;

    if (listaTareas.length === 0) {
      registro.__estado_proceso = "completado_ok";
      registro.__progreso.mensajeActual = "No hay tareas pendientes. Todo parece estar sincronizado.";
      const nombreArchivoRegistroFinal = generarNombreRegistro(nombreDestino);
      guardarRegistro(registro, nombreArchivoRegistroFinal);
      PropertiesService.getUserProperties().setProperty(PS_KEYS.LOG_FILENAME, nombreArchivoRegistroFinal);
      limpiarPropiedadesBatch(); // No hay nada que hacer por lotes

      return {
        mensaje: `✅ ${registro.__progreso.mensajeActual}`,
        idDestino: idCarpetaDestino,
        nombreDestino: nombreDestino,
        urlRegistro: DriveApp.getFileById(guardarRegistro(registro, nombreArchivoRegistroFinal).getId()).getUrl(), // Asegúrate que guardarRegistro devuelve el File
        procesoPorLotesIniciado: false
      };
    }

    registro.__estado_proceso = "en_progreso";
    const nombreArchivoRegistro = generarNombreRegistro(nombreDestino); // Tu función existente
    const archivoRegistroGuardado = guardarRegistro(registro, nombreArchivoRegistro); // Asumimos que guardarRegistro devuelve el objeto File
  

    // Configurar PropertiesService para el proceso por lotes
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperties({
      [PS_KEYS.BATCH_ACTIVE]: 'true',
      [PS_KEYS.LOG_FILENAME]: nombreArchivoRegistro,
      [PS_KEYS.TASK_INDEX]: '0'
    });

    // Crear el primer trigger
    crearSiguienteTrigger('procesarLoteTareas_triggered');

    return {
      mensaje: `🚀 Proceso de copia por lotes iniciado para "${nombreDestino}". ${listaTareas.length} tareas en cola. El proceso continuará en segundo plano.`,
      idDestino: idCarpetaDestino,
      nombreDestino: nombreDestino,
      urlRegistro: archivoRegistroGuardado ? archivoRegistroGuardado.getUrl() : null,
      procesoPorLotesIniciado: true
    };

  } catch (e) {
    Logger.log("Error general en iniciarCopiaPorLotes: " + e.stack);
    // Limpiar propiedades en caso de error en la inicialización
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    return { 
        mensaje: `❌ Ocurrió un error inesperado al iniciar: ${e.message}. Revisa los registros del script.`,
        procesoPorLotesIniciado: false
    };
  }
}

// ==========================================================================
// FUNCIONES PARA EL PROCESAMIENTO POR LOTES (TRIGGERS) 
// ==========================================================================

/**
 * Función que es llamada directamente por el trigger.
 * Actúa como un envoltorio seguro para la lógica principal del lote.
 */
function procesarLoteTareas_triggered() {
  const lock = LockService.getUserLock();
  if (lock.tryLock(30000)) { // Intentar obtener un bloqueo durante 30s
    try {
      Logger.log("Procesando lote de tareas (triggered)...");
      procesarLoteTareas();
    } catch (e) {
      Logger.log(`Error crítico en procesarLoteTareas_triggered: ${e.stack}`);
      const userProps = PropertiesService.getUserProperties();
      const logFileName = userProps.getProperty(PS_KEYS.LOG_FILENAME);
      if (logFileName) {
        try {
            let registro = JSON.parse(obtenerOCrearCarpeta(NOMBRE_CARPETA_INFORMES).getFilesByName(logFileName).next().getBlob().getDataAsString());
            registro.__estado_proceso = "error_critico";
            registro.__errores.push({timestamp: new Date().toISOString(), tarea: "CRITICO", mensaje: `Error en trigger: ${e.message}`});
            registro.__progreso.mensajeActual = "Error crítico en el proceso por lotes.";
            guardarRegistro(registro, logFileName); // Asumiendo que guardarRegistro maneja File object o crea uno nuevo
        } catch (logError) {
            Logger.log(`Error al intentar guardar el error crítico en el log: ${logError.stack}`);
        }
      }
      borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
      limpiarPropiedadesBatch();
    } finally {
      lock.releaseLock();
    }
  } else {
    Logger.log("No se pudo obtener el bloqueo para procesarLoteTareas_triggered. Otro proceso podría estar en ejecución.");
  }
}


/**
 * Lógica principal para procesar un lote de tareas. VERSIÓN ACTUALIZADA CON EJECUCIÓN DE TAREAS.
 */
function procesarLoteTareas() {
  const userProps = PropertiesService.getUserProperties();
  if (userProps.getProperty(PS_KEYS.BATCH_ACTIVE) !== 'true') {
    Logger.log("Proceso por lotes no activo. Deteniendo.");
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    return;
  }

  const logFileName = userProps.getProperty(PS_KEYS.LOG_FILENAME);
  // Este es el índice de la primera tarea que este lote intentará procesar.
  let globalTaskPointer = parseInt(userProps.getProperty(PS_KEYS.TASK_INDEX) || '0');
  const idCarpetaDestinoRaiz = userProps.getProperty(PS_KEYS.DESTINATION_FOLDER_ID); // Usado si rutaPadreDestinoLogKey es la raíz

  if (!logFileName || !idCarpetaDestinoRaiz) {
    Logger.log("Faltan propiedades esenciales (LOG_FILENAME o DESTINATION_FOLDER_ID). Deteniendo proceso.");
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    return;
  }
  
  let registro;
  let archivoLog; // DriveApp.File object
  try {
    const carpetaInformes = obtenerOCrearCarpeta(NOMBRE_CARPETA_INFORMES);
    const iteradorArchivos = carpetaInformes.getFilesByName(logFileName);
    if (!iteradorArchivos.hasNext()) {
        Logger.log(`Archivo de log ${logFileName} no encontrado. Deteniendo.`);
        limpiarPropiedadesBatch();
        borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
        return;
    }
    archivoLog = iteradorArchivos.next();
    registro = JSON.parse(archivoLog.getBlob().getDataAsString());
  } catch (e) {
    Logger.log(`Error al cargar el archivo de log ${logFileName}: ${e.stack}. Deteniendo.`);
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    return;
  }

  // Si el estado ya es final, no hacer nada más y limpiar.
  if (registro.__estado_proceso === "completado_ok" || registro.__estado_proceso === "completado_con_errores" || registro.__estado_proceso === "error_critico") {
    Logger.log(`El proceso ya está marcado como '${registro.__estado_proceso}'. Limpiando y saliendo.`);
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    return;
  }
  
  if (!registro.__pendientes || globalTaskPointer >= registro.__pendientes.length) {
    Logger.log("No hay más tareas pendientes o el índice es inválido. Marcando como completado.");
    registro.__estado_proceso = registro.__errores?.length > 0 ? "completado_con_errores" : "completado_ok";
    registro.__progreso.mensajeActual = "Todas las tareas han sido procesadas.";
    registro.__fecha = new Date().toISOString();
    // Guardar estado final (usando guardarRegistro para consistencia)
    guardarRegistro(registro, logFileName); 
    
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    return;
  }

  registro.__progreso.erroresEnLote = 0; // Resetear contador de errores del lote actual

  // Procesar hasta LOTE_TAMAÑO tareas o hasta que una tarea falle y necesite reintento sin avanzar el puntero.
  for (let i = 0; i < LOTE_TAMAÑO && globalTaskPointer < registro.__pendientes.length; i++) {
    // NO incrementamos i aquí, solo si la tarea se "resuelve" (éxito o fallo permanente)
    // o si la tarea actual falla pero podemos pasar a la siguiente del lote.
    // La variable 'i' cuenta cuántas tareas del LOTE_TAMAÑO hemos *intentado*.
    // globalTaskPointer es el que realmente importa para el progreso general.

    const tarea = registro.__pendientes[globalTaskPointer];
    // Podríamos añadir un estado a la tarea misma, ej. tarea.estado = 'procesando'
    
    registro.__progreso.mensajeActual = `Procesando tarea ${globalTaskPointer + 1}/${registro.__progreso.totalTareas}: ${tarea.tipo} - ${tarea.nombre}`;
    Logger.log(registro.__progreso.mensajeActual);

    let tareaResueltaEnEsteIntento = false; // Indica si la tarea actual avanzó (éxito o fallo max)

    try {
      if (tarea.tipo === "crearCarpeta") {
        ejecutarTareaCrearCarpeta(tarea, registro, idCarpetaDestinoRaiz);
      } else if (tarea.tipo === "copiarArchivo") {
        ejecutarTareaCopiarArchivo(tarea, registro, idCarpetaDestinoRaiz);
      } else {
        throw new Error(`Tipo de tarea desconocido: ${tarea.tipo}`);
      }
      
      // Si llegamos aquí, la tarea fue exitosa
      registro.__progreso.tareasCompletadas++;
      tarea.estado = "completada_ok"; // Marcar estado en la tarea misma (opcional)
      Logger.log(`Tarea '${tarea.nombre}' completada exitosamente.`);
      tareaResueltaEnEsteIntento = true;

    } catch (e) {
      const mensajeError = `Error procesando tarea '${tarea.nombre}' (Intento ${ (tarea.intentos || 0) + 1}): ${e.message}`;
      Logger.log(mensajeError);
      Logger.log(e.stack); // Loguear el stacktrace completo para depuración
      
      registro.__errores.push({timestamp: new Date().toISOString(), tarea: tarea, mensaje: e.message});
      registro.__progreso.erroresEnLote++;
      tarea.intentos = (tarea.intentos || 0) + 1;

      if (tarea.intentos >= MAX_INTENTOS_POR_TAREA) {
        registro.__progreso.mensajeActual = `Tarea '${tarea.nombre}' falló ${MAX_INTENTOS_POR_TAREA} veces. Se omite permanentemente.`;
        Logger.log(registro.__progreso.mensajeActual);
        // Contar como "completada" para que el progreso general avance y el proceso pueda terminar.
        registro.__progreso.tareasCompletadas++; 
        tarea.estado = "fallida_permanente"; // Marcar estado en la tarea misma (opcional)
        tareaResueltaEnEsteIntento = true; // Se considera resuelta (aunque fallida)
      } else {
        registro.__progreso.mensajeActual = `Tarea '${tarea.nombre}' falló, se reintentará.`;
        Logger.log(registro.__progreso.mensajeActual);
        // No se considera resuelta, globalTaskPointer NO avanzará.
        // El lote actual se interrumpe aquí para esta tarea, para darle prioridad en el siguiente trigger.
        tareaResueltaEnEsteIntento = false; 
        // Rompemos el bucle for para que este trigger guarde el estado actual y
        // el siguiente trigger reintente esta misma tarea primero.
        break; 
      }
    }

    if (tareaResueltaEnEsteIntento) {
      globalTaskPointer++; // Avanzar al siguiente índice de tarea solo si la actual se resolvió
    }
  } // Fin del bucle for que procesa el LOTE_TAMAÑO
  
  userProps.setProperty(PS_KEYS.TASK_INDEX, globalTaskPointer.toString());
  registro.__fecha = new Date().toISOString();

  try {
    // Usar guardarRegistro para asegurar consistencia si esa función maneja la creación/actualización de File object
    // Si tu guardarRegistro espera el objeto JSON y el nombre, y devuelve el File object:
    const archivoLogActualizado = guardarRegistro(registro, logFileName); 
    if (!archivoLogActualizado) { // guardarRegistro debería devolver el objeto File o null/error
        throw new Error("guardarRegistro no devolvió un objeto de archivo válido.");
    }
    // Si necesitas el contenido actualizado para algo más inmediatamente:
    // archivoLog.setContent(Utilities.newBlob(JSON.stringify(registro, null, 2), MimeType.PLAIN_TEXT, logFileName).getDataAsString());
  } catch(e) {
      Logger.log(`FALLO CRITICO: No se pudo guardar el archivo de log ${logFileName}. ${e.stack}`);
      limpiarPropiedadesBatch();
      borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
      return; 
  }

  if (globalTaskPointer < registro.__pendientes.length && registro.__progreso.tareasCompletadas < registro.__progreso.totalTareas) {
    crearSiguienteTrigger('procesarLoteTareas_triggered');
    Logger.log(`Lote procesado. Puntero global de tareas en: ${globalTaskPointer}. Próximo trigger programado.`);
  } else {
    Logger.log("Todas las tareas han sido procesadas o han alcanzado el límite de reintentos. Limpiando.");
    registro.__estado_proceso = registro.__errores?.length > 0 ? "completado_con_errores" : "completado_ok";
    if (registro.__progreso.tareasCompletadas >= registro.__progreso.totalTareas) {
         registro.__progreso.mensajeActual = `Proceso finalizado. Tareas procesadas: ${registro.__progreso.tareasCompletadas} de ${registro.__progreso.totalTareas}.`;
    } else {
         registro.__progreso.mensajeActual = `Proceso finalizado con tareas pendientes no resueltas. Procesadas: ${registro.__progreso.tareasCompletadas} de ${registro.__progreso.totalTareas}.`;
    }
    registro.__fecha = new Date().toISOString();
    guardarRegistro(registro, logFileName);
    
    limpiarPropiedadesBatch();
    borrarTriggerActualPorNombre('procesarLoteTareas_triggered');
    Logger.log("Proceso por lotes finalizado y triggers eliminados.");
  }
}


// ==========================================================================
// FUNCIONES AUXILIARES PARA EJECUTAR TAREAS INDIVIDUALES
// ==========================================================================

/**
 * Ejecuta una tarea de tipo "crearCarpeta".
 * Actualiza el objeto 'registro' con el ID de la carpeta creada.
 * @param {Object} tarea Objeto de la tarea a ejecutar.
 * @param {Object} registro Objeto de log principal (se modifica directamente).
 * @param {String} idCarpetaDestinoRaiz ID de la carpeta destino raíz (por si la ruta padre es la raíz).
 */
function ejecutarTareaCrearCarpeta(tarea, registro, idCarpetaDestinoRaiz) {
  Logger.log(`Ejecutando crearCarpeta: ${tarea.nombre} en ${tarea.rutaPadreDestinoLogKey}`);
  let idCarpetaPadreDestino;

  if (registro[tarea.rutaPadreDestinoLogKey] && registro[tarea.rutaPadreDestinoLogKey].id_destino_folder) {
    idCarpetaPadreDestino = registro[tarea.rutaPadreDestinoLogKey].id_destino_folder;
  } else {
    // Esto podría pasar si la ruta padre es la raíz misma y no tiene una entrada '.id_destino_folder' separada,
    // o si hay un error en la lógica de planificación/registro.
    // Comprobar si la ruta padre es la raíz del log.
    if (tarea.rutaPadreDestinoLogKey === `/${registro.__raiz}`) {
        idCarpetaPadreDestino = idCarpetaDestinoRaiz;
    } else {
        throw new Error(`No se pudo encontrar el ID de la carpeta destino padre '${tarea.rutaPadreDestinoLogKey}' en el registro.`);
    }
  }
  
  const carpetaPadreDestino = DriveApp.getFolderById(idCarpetaPadreDestino);
  
  // Verificar si ya existe una carpeta con ese nombre (podría haber sido creada por otro proceso o manualmente)
  const carpetasExistentes = carpetaPadreDestino.getFoldersByName(tarea.nombre);
  let nuevaCarpeta;
  if (carpetasExistentes.hasNext()) {
      nuevaCarpeta = carpetasExistentes.next();
      Logger.log(`Carpeta '${tarea.nombre}' ya existía en destino. Usando existente.`);
  } else {
      nuevaCarpeta = carpetaPadreDestino.createFolder(tarea.nombre);
      Logger.log(`Carpeta '${tarea.nombre}' creada con ID: ${nuevaCarpeta.getId()}`);
  }
  
  const idNuevaCarpeta = nuevaCarpeta.getId();

  // Actualizar el registro:
  // 1. En la entrada del padre, registrar el ID de esta nueva subcarpeta.
  if (!registro[tarea.rutaPadreDestinoLogKey].carpetas) {
    registro[tarea.rutaPadreDestinoLogKey].carpetas = {};
  }
  registro[tarea.rutaPadreDestinoLogKey].carpetas[tarea.nombre] = idNuevaCarpeta;

  // 2. En la entrada propia de la nueva carpeta (rutaOrigenAbsoluta), registrar su ID de destino.
  if (!registro[tarea.rutaOrigenAbsoluta]) {
    registro[tarea.rutaOrigenAbsoluta] = { archivos: {}, carpetas: {} };
  }
  registro[tarea.rutaOrigenAbsoluta].id_destino_folder = idNuevaCarpeta;
}

/**
 * Ejecuta una tarea de tipo "copiarArchivo".
 * Actualiza el objeto 'registro' con el ID y metadatos del archivo copiado.
 * @param {Object} tarea Objeto de la tarea a ejecutar.
 * @param {Object} registro Objeto de log principal (se modifica directamente).
 * @param {String} idCarpetaDestinoRaiz ID de la carpeta destino raíz.
 */
function ejecutarTareaCopiarArchivo(tarea, registro, idCarpetaDestinoRaiz) {
  Logger.log(`Ejecutando copiarArchivo: ${tarea.nombre} a ${tarea.rutaPadreDestinoLogKey}`);
  let idCarpetaPadreDestino;

  if (registro[tarea.rutaPadreDestinoLogKey] && registro[tarea.rutaPadreDestinoLogKey].id_destino_folder) {
    idCarpetaPadreDestino = registro[tarea.rutaPadreDestinoLogKey].id_destino_folder;
  } else {
     if (tarea.rutaPadreDestinoLogKey === `/${registro.__raiz}`) {
        idCarpetaPadreDestino = idCarpetaDestinoRaiz;
    } else {
        throw new Error(`No se pudo encontrar el ID de la carpeta destino padre '${tarea.rutaPadreDestinoLogKey}' para el archivo '${tarea.nombre}'.`);
    }
  }

  const carpetaPadreDestinoDrive = DriveApp.getFolderById(idCarpetaPadreDestino);
  const archivoOrigen = DriveApp.getFileById(tarea.idArchivoOrigen);

  // Antes de copiar, verificar si un archivo con el mismo nombre ya existe en destino y borrarlo.
  // Esto asegura que la copia sea "fresca" y evita el "(1)" de Drive si se hace makeCopy sobre existente.
  const archivosExistentes = carpetaPadreDestinoDrive.getFilesByName(tarea.nombre);
  while (archivosExistentes.hasNext()) {
    const archivoExistente = archivosExistentes.next();
    // Comprobar si es el mismo que está registrado (si aplica)
    const archivoRegistrado = registro[tarea.rutaPadreDestinoLogKey]?.archivos?.[tarea.nombre];
    if (archivoRegistrado && archivoRegistrado.id === archivoExistente.getId()) {
        // Es el archivo que registramos previamente, probablemente modificado.
        Logger.log(`Archivo '${tarea.nombre}' (ID: ${archivoExistente.getId()}) existe y está registrado. Se reemplazará (borrando y copiando).`);
    } else {
        Logger.log(`Archivo con nombre '${tarea.nombre}' (ID: ${archivoExistente.getId()}) encontrado en destino. Se reemplazará (borrando y copiando).`);
    }
    archivoExistente.setTrashed(true); // Enviar a la papelera
  }
  
  const copiaArchivo = archivoOrigen.makeCopy(tarea.nombre, carpetaPadreDestinoDrive);
  Logger.log(`Archivo '${tarea.nombre}' copiado con ID: ${copiaArchivo.getId()}`);
  
  // Actualizar el registro:
  if (!registro[tarea.rutaPadreDestinoLogKey].archivos) {
    registro[tarea.rutaPadreDestinoLogKey].archivos = {};
  }
  registro[tarea.rutaPadreDestinoLogKey].archivos[tarea.nombre] = {
    id: copiaArchivo.getId(),
    size: tarea.sizeOrigen,
    modificado: tarea.modificadoOrigen // Guardar metadatos del origen
  };
}

// ==========================================================================
// FUNCIÓN DE PLANIFICACIÓN (DETALLADA)
// ==========================================================================
/**
 * Genera la lista de tareas pendientes (archivos a copiar, carpetas a crear).
 * @param {GoogleAppsScript.Drive.Folder} carpetaOrigenRaiz El objeto Folder de origen raíz.
 * @param {GoogleAppsScript.Drive.Folder} carpetaDestinoRaiz El objeto Folder de destino raíz.
 * @param {String} nombreRaizOrigen El nombre de la carpeta raíz de origen (para las rutas en el log).
 * @param {Object} registro El objeto de log actual (puede tener información de ejecuciones previas).
 * @param {Array} listaTareas Array (pasado por referencia) donde se añadirán las tareas.
 */
function generarListaDeTareas(carpetaOrigenRaiz, carpetaDestinoRaiz, nombreRaizOrigen, registro, listaTareas) {
  Logger.log(`Iniciando generación de lista de tareas para origen: ${carpetaOrigenRaiz.getName()}, destino: ${carpetaDestinoRaiz.getName()}`);
  const rutaLogRaiz = `/${nombreRaizOrigen}`;

  // Asegurar que la ruta raíz existe en el registro y tiene el ID de la carpeta destino raíz.
  if (!registro[rutaLogRaiz]) {
    registro[rutaLogRaiz] = { archivos: {}, carpetas: {} };
  }
  // Este ID es el punto de partida para todas las operaciones en el destino.
  registro[rutaLogRaiz].id_destino_folder = carpetaDestinoRaiz.getId();

  // Iniciar la planificación recursiva desde la raíz.
  planificarRecursivamente(carpetaOrigenRaiz, rutaLogRaiz, registro, listaTareas, carpetaDestinoRaiz);
  
  Logger.log(`Generación de lista de tareas finalizada. Total tareas planificadas: ${listaTareas.length}`);
}

/**
 * Función auxiliar recursiva para planificar tareas.
 * @param {GoogleAppsScript.Drive.Folder} carpetaOrigenActual Carpeta origen que se está procesando.
 * @param {String} rutaLogActual Key en el 'registro' para la carpetaOrigenActual (ej: "/Raiz/SubDir").
 * @param {Object} registro El objeto de log principal.
 * @param {Array} listaTareas Array donde se añaden las tareas.
 * @param {GoogleAppsScript.Drive.Folder | null} carpetaDestinoDriveActual Objeto Folder de Drive para la carpeta destino correspondiente.
 * Puede ser 'null' si la carpeta destino aún no existe (y está por ser creada).
 */
function planificarRecursivamente(carpetaOrigenActual, rutaLogActual, registro, listaTareas, carpetaDestinoDriveActual) {
  Logger.log(`Planificando para: ${rutaLogActual}`);

  // Asegurar que la entrada para la ruta actual exista en el registro
  if (!registro[rutaLogActual]) {
    registro[rutaLogActual] = { archivos: {}, carpetas: {} };
  }
  // Si tenemos el objeto carpetaDestinoDriveActual, aseguramos que su ID esté en el registro.
  if (carpetaDestinoDriveActual) {
      registro[rutaLogActual].id_destino_folder = carpetaDestinoDriveActual.getId();
  }


  // --- 1. PROCESAR SUBCARPETAS ---
  const subcarpetasOrigen = carpetaOrigenActual.getFolders();
  while (subcarpetasOrigen.hasNext()) {
    const subCarpetaO = subcarpetasOrigen.next();
    const nombreSubCarpeta = subCarpetaO.getName();
    const rutaLogSubCarpeta = `${rutaLogActual}/${nombreSubCarpeta}`; // Path para el registro

    let subCarpetaDestinoDrive = null; // Objeto Folder para la subcarpeta en destino
    let necesitaCrearTareaCarpeta = true;

    // Verificar contra el registro y Drive
    const idSubCarpetaDestinoRegistrada = registro[rutaLogActual]?.carpetas?.[nombreSubCarpeta];
    if (idSubCarpetaDestinoRegistrada) {
      try {
        const folder = DriveApp.getFolderById(idSubCarpetaDestinoRegistrada);
        if (folder.isTrashed()) {
          Logger.log(`Carpeta destino '${rutaLogSubCarpeta}' (ID: ${idSubCarpetaDestinoRegistrada}) está en papelera. Se marcará para recrear.`);
          delete registro[rutaLogActual].carpetas[nombreSubCarpeta]; // Limpiar registro de ID inválido
          if(registro[rutaLogSubCarpeta]) delete registro[rutaLogSubCarpeta].id_destino_folder;
        } else {
          subCarpetaDestinoDrive = folder;
          necesitaCrearTareaCarpeta = false;
        }
      } catch (e) {
        Logger.log(`Carpeta destino '${rutaLogSubCarpeta}' (ID: ${idSubCarpetaDestinoRegistrada}) no accesible: ${e.message}. Se marcará para recrear.`);
        delete registro[rutaLogActual].carpetas[nombreSubCarpeta]; // Limpiar registro de ID inválido
        if(registro[rutaLogSubCarpeta]) delete registro[rutaLogSubCarpeta].id_destino_folder;
      }
    } else if (carpetaDestinoDriveActual) { // Si no está en registro, pero la carpeta padre destino existe, buscar por nombre
      const iteradorCarpetasDestino = carpetaDestinoDriveActual.getFoldersByName(nombreSubCarpeta);
      if (iteradorCarpetasDestino.hasNext()) {
        subCarpetaDestinoDrive = iteradorCarpetasDestino.next();
        // Encontramos una carpeta existente no registrada. La registramos.
        if (!registro[rutaLogActual].carpetas) registro[rutaLogActual].carpetas = {};
        registro[rutaLogActual].carpetas[nombreSubCarpeta] = subCarpetaDestinoDrive.getId();
        if (!registro[rutaLogSubCarpeta]) registro[rutaLogSubCarpeta] = { archivos: {}, carpetas: {} };
        registro[rutaLogSubCarpeta].id_destino_folder = subCarpetaDestinoDrive.getId();
        necesitaCrearTareaCarpeta = false;
        Logger.log(`Subcarpeta '${nombreSubCarpeta}' encontrada en destino (no registrada previamente). Registrada.`);
      }
    }

    if (necesitaCrearTareaCarpeta) {
      listaTareas.push({
        tipo: "crearCarpeta",
        nombre: nombreSubCarpeta,
        rutaOrigenAbsoluta: rutaLogSubCarpeta,
        rutaPadreDestinoLogKey: rutaLogActual, // El padre es la carpeta actual que estamos procesando
        intentos: 0
      });
      // Preparamos la entrada en el registro para esta nueva carpeta, aunque aún no tenga ID de destino.
      if (!registro[rutaLogSubCarpeta]) registro[rutaLogSubCarpeta] = { archivos: {}, carpetas: {} };
      // Llamada recursiva: pasamos 'null' como carpetaDestinoDrive, ya que aún no se ha creado.
      planificarRecursivamente(subCarpetaO, rutaLogSubCarpeta, registro, listaTareas, null);
    } else if (subCarpetaDestinoDrive) { // La carpeta existe en destino
      // Asegurar que el ID de destino esté en la entrada propia de la subcarpeta en el registro.
      if (!registro[rutaLogSubCarpeta]) registro[rutaLogSubCarpeta] = { archivos: {}, carpetas: {} };
      registro[rutaLogSubCarpeta].id_destino_folder = subCarpetaDestinoDrive.getId();
      planificarRecursivamente(subCarpetaO, rutaLogSubCarpeta, registro, listaTareas, subCarpetaDestinoDrive);
    }
  }

  // --- 2. PROCESAR ARCHIVOS ---
  const archivosOrigen = carpetaOrigenActual.getFiles();
  while (archivosOrigen.hasNext()) {
    const archivoO = archivosOrigen.next();
    const nombreArchivo = archivoO.getName();
    // La rutaLogArchivo es la rutaOrigenAbsoluta para la tarea
    const rutaLogArchivo = `${rutaLogActual}/${nombreArchivo}`;

    const sizeO = archivoO.getSize();
    const modificadoO = archivoO.getLastUpdated().getTime();
    let necesitaCrearTareaArchivo = true;

    // Verificar contra el registro
    const infoArchivoRegistrado = registro[rutaLogActual]?.archivos?.[nombreArchivo];
    if (infoArchivoRegistrado) {
      if (infoArchivoRegistrado.size === sizeO && infoArchivoRegistrado.modificado === modificadoO) {
        try {
          const archivoDestinoDrive = DriveApp.getFileById(infoArchivoRegistrado.id);
          if (!archivoDestinoDrive.isTrashed()) {
            necesitaCrearTareaArchivo = false; // Coincide y existe
          } else {
            Logger.log(`Archivo destino '${rutaLogArchivo}' (ID: ${infoArchivoRegistrado.id}) en papelera. Se marcará para recopiar.`);
            // No es necesario borrar del registro aquí, la tarea de copia lo sobrescribirá.
          }
        } catch (e) {
          Logger.log(`Archivo destino '${rutaLogArchivo}' (ID: ${infoArchivoRegistrado.id}) no accesible: ${e.message}. Se marcará para recopiar.`);
        }
      } else {
         Logger.log(`Archivo '${rutaLogArchivo}' ha cambiado (tamaño/fecha) desde el último registro. Se marcará para recopiar.`);
      }
    } else if (carpetaDestinoDriveActual) { // Si no está en registro, pero la carpeta padre destino existe, buscar por nombre
        const iteradorArchivosDestino = carpetaDestinoDriveActual.getFilesByName(nombreArchivo);
        if (iteradorArchivosDestino.hasNext()) {
            const archivoDestinoDrive = iteradorArchivosDestino.next();
            // Comparamos solo tamaño por simplicidad (fechas de modificación pueden variar con copias)
            // Para una mayor precisión, se podría comparar un hash si fuera viable.
            if (archivoDestinoDrive.getSize() === sizeO) {
                // Encontramos un archivo no registrado que parece ser el mismo. Lo registramos.
                if (!registro[rutaLogActual].archivos) registro[rutaLogActual].archivos = {};
                registro[rutaLogActual].archivos[nombreArchivo] = {
                    id: archivoDestinoDrive.getId(),
                    size: sizeO,
                    modificado: modificadoO // Guardamos la fecha del origen
                };
                necesitaCrearTareaArchivo = false;
                Logger.log(`Archivo '${nombreArchivo}' encontrado en destino (no registrado previamente) y coincide en tamaño. Registrado.`);
            } else {
                 Logger.log(`Archivo '${nombreArchivo}' encontrado en destino (no registrado previamente) pero difiere en tamaño. Se marcará para copiar (sobrescribir).`);
            }
        }
    }


    if (necesitaCrearTareaArchivo) {
      listaTareas.push({
        tipo: "copiarArchivo",
        nombre: nombreArchivo,
        idArchivoOrigen: archivoO.getId(),
        rutaOrigenAbsoluta: rutaLogArchivo,
        rutaPadreDestinoLogKey: rutaLogActual,
        sizeOrigen: sizeO,
        modificadoOrigen: modificadoO,
        intentos: 0
      });
      // Registrar archivos grandes (opcional, ya lo tenías)
      if (sizeO > 100 * 1024 * 1024 && registro.__pesados && !registro.__pesados.some(p => p.startsWith(rutaLogArchivo))) {
         registro.__pesados.push(`${rutaLogArchivo} (${Math.round(sizeO / 1024 / 1024)} MB)`);
      }
    }
  }
}


// ==========================================================================
// FUNCIONES DE UTILIDAD PARA TRIGGERS Y PROPERTIES
// ==========================================================================
function crearSiguienteTrigger(nombreFuncionTrigger) {
  // Primero, borrar cualquier trigger existente con el mismo manejador para evitar duplicados.
  // Esto es crucial si el script se detuvo inesperadamente antes de borrar su propio trigger.
  borrarTriggerActualPorNombre(nombreFuncionTrigger); 

  const trigger = ScriptApp.newTrigger(nombreFuncionTrigger)
    .timeBased()
    .after(TRIGGER_DELAY_MS)
    .create();
  PropertiesService.getUserProperties().setProperty(PS_KEYS.TRIGGER_ID, trigger.getUniqueId());
  Logger.log(`Trigger creado con ID: ${trigger.getUniqueId()} para ejecutar ${nombreFuncionTrigger} en ${TRIGGER_DELAY_MS / 1000}s.`);
}

function borrarTriggerActualPorNombre(nombreFuncionTrigger) {
  const userProps = PropertiesService.getUserProperties();
  const triggerIdActual = userProps.getProperty(PS_KEYS.TRIGGER_ID);

  const projectTriggers = ScriptApp.getProjectTriggers();
  let borrado = false;

  // Intenta borrar por ID si lo tenemos
  if (triggerIdActual) {
    for (let i = 0; i < projectTriggers.length; i++) {
      if (projectTriggers[i].getUniqueId() === triggerIdActual) {
        ScriptApp.deleteTrigger(projectTriggers[i]);
        Logger.log(`Trigger con ID ${triggerIdActual} borrado.`);
        borrado = true;
        break;
      }
    }
  }
  
  // Si no se borró por ID (o no teníamos ID), intentar borrar por nombre de función (menos preciso)
  // Esto sirve como limpieza adicional si el ID se perdió o si hay triggers huérfanos.
  if (!borrado) {
    for (let i = 0; i < projectTriggers.length; i++) {
      if (projectTriggers[i].getHandlerFunction() === nombreFuncionTrigger) {
        ScriptApp.deleteTrigger(projectTriggers[i]);
        Logger.log(`Trigger huérfano para la función ${nombreFuncionTrigger} borrado.`);
        // No rompemos el bucle aquí, podría haber múltiples huérfanos si algo salió muy mal.
      }
    }
  }
  userProps.deleteProperty(PS_KEYS.TRIGGER_ID); // Limpiar el ID de la propiedad
}

function limpiarPropiedadesBatch() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty(PS_KEYS.BATCH_ACTIVE);
  userProps.deleteProperty(PS_KEYS.LOG_FILENAME);
  userProps.deleteProperty(PS_KEYS.TASK_INDEX);
  userProps.deleteProperty(PS_KEYS.TRIGGER_ID); // Asegurarse de limpiar el ID del trigger
  userProps.deleteProperty(PS_KEYS.DESTINATION_FOLDER_ID);
  userProps.deleteProperty(PS_KEYS.SOURCE_FOLDER_ID);
  Logger.log("Propiedades de usuario para el proceso por lotes limpiadas.");
}


// ==========================================================================
// FUNCIÓN PARA LA INTERFAZ HTML (SONDEO DE ESTADO)
// ==========================================================================
function obtenerEstadoCopia() {
  const userProps = PropertiesService.getUserProperties();
  const batchActivo = userProps.getProperty(PS_KEYS.BATCH_ACTIVE) === 'true';
  const logFileName = userProps.getProperty(PS_KEYS.LOG_FILENAME);
  const idDestinoActual = userProps.getProperty(PS_KEYS.DESTINATION_FOLDER_ID); // Obtenerlo siempre


  if (!batchActivo && !logFileName) { // Ni activo, ni un log reciente de un proceso
    return {
      procesoActivo: false,
      mensaje: "No hay ningún proceso de copia activo o información de un proceso reciente.",
      progreso: null,
      urlRegistro: null
    };
  }
  
  // Si no está activo pero hay un logFileName, significa que el proceso terminó.
  // Cargamos ese último log para dar el estado final.
  if (!logFileName) { // Debería haber sido limpiado si no está activo, pero por si acaso.
      return { procesoActivo: batchActivo, mensaje: "Proceso activo pero sin archivo de log configurado (estado inconsistente).", progreso: null, urlRegistro: null };
  }

  try {
    const carpetaInformes = obtenerOCrearCarpeta(NOMBRE_CARPETA_INFORMES);
    const archivos = carpetaInformes.getFilesByName(logFileName);
    let urlReg = null;
    if (archivos.hasNext()) {
      const archivoLog = archivos.next();
      const registro = JSON.parse(archivoLog.getBlob().getDataAsString());
      urlReg = archivoLog.getUrl();
      return {
        procesoActivo: batchActivo, // Podría ser true (en curso) o false (recién terminado)
        mensaje: registro.__progreso?.mensajeActual || "Consultando estado...",
        progreso: registro.__progreso,
        estadoGeneral: registro.__estado_proceso,
        totalTareas: registro.__progreso?.totalTareas || 0,
        tareasCompletadas: registro.__progreso?.tareasCompletadas || 0,
        errores: registro.__errores,
        urlRegistro: archivoLog.getUrl()
      };
    } else {
      // Si el log no se encuentra pero se esperaba, es un problema.
      if (batchActivo) {
          return { 
            procesoActivo: batchActivo,
        mensaje: registro.__progreso?.mensajeActual || (batchActivo ? "Consultando estado..." : "Proceso finalizado."),
        progreso: registro.__progreso,
        estadoGeneral: registro.__estado_proceso,
        totalTareas: registro.__progreso?.totalTareas || 0,
        tareasCompletadas: registro.__progreso?.tareasCompletadas || 0,
        errores: registro.__errores,
        urlRegistro: urlReg, // Devolver siempre la URL del log actual si existe
        idDestino: idDestinoActual // Devolver siempre el ID del destino si existe
          };
      } else { // Proceso no activo, y el log que teníamos registrado ya no está.
          return { procesoActivo: false, mensaje: "No se encontró información del último proceso.", progreso: null, urlRegistro: null };
      }
    }
  } catch (e) {
    Logger.log("Error en obtenerEstadoCopia: " + e.stack);
    return {
      procesoActivo: batchActivo, // Devuelve el estado de actividad conocido
      mensaje: "Error al obtener el estado de la copia: " + e.message,
      progreso: null,
      urlRegistro: null
    };
  }
}


// ==========================================================================
// TUS FUNCIONES EXISTENTES (revisar y adaptar si es necesario más adelante)
// ==========================================================================

function verificarIntegridad(datos) {
  let idCarpetaDestino = null;
  let nombreDestinoReal = datos.nombreDestino ? datos.nombreDestino.trim() : null;
  let urlRegistro = null;

  try {
    const idOrigen = datos.idCarpeta.trim();
    if (!nombreDestinoReal) throw new Error("El nombre de la carpeta destino es requerido.");

    const carpetaOrigen = DriveApp.getFolderById(idOrigen);
    const carpetasDestinoIter = DriveApp.getFoldersByName(nombreDestinoReal);

    if (!carpetasDestinoIter.hasNext()) {
      return { 
        mensaje: `❌ No se encontró la carpeta de destino "${nombreDestinoReal}".`, 
        esCoincidente: false, 
        idDestino: null, 
        urlRegistro: null,
        nombreDestino: nombreDestinoReal
      };
    }
    const carpetaDestino = carpetasDestinoIter.next();
    idCarpetaDestino = carpetaDestino.getId();

    const dataRegistro = cargarRegistro(nombreDestinoReal);
    const registro = dataRegistro.registroJson;
    if (dataRegistro.archivoLog) {
      urlRegistro = dataRegistro.archivoLog.getUrl();
    }

    if (Object.keys(registro).length === 0 || !registro.__raiz) { // Si el registro está vacío o no tiene __raiz
        return {
            mensaje: `ℹ️ No existe un registro de copia válido o completo para "${nombreDestinoReal}". No se puede verificar contra él. Considere ejecutar una copia primero.`,
            esCoincidente: false, 
            idDestino: idCarpetaDestino,
            urlRegistro: urlRegistro, // Puede haber un log vacío, pero lo pasamos.
            nombreDestino: nombreDestinoReal
        };
    }

    const nombreRaizOrigen = registro.__raiz || carpetaOrigen.getName();
    const resultadoComparacion = compararEstructuras(carpetaOrigen, carpetaDestino, "", registro, nombreRaizOrigen);

    let mensajeFinal;
    let esCoincidenteLocal = false;

    if (resultadoComparacion.trim() === "") {
      mensajeFinal = "✅ Todo coincide: no se encontraron diferencias entre origen y destino según el último registro.";
      esCoincidenteLocal = true;
    } else {
      mensajeFinal = `
        <strong>⚠ Verificación finalizada con incidencias (comparado con origen y último registro):</strong><br><br>${resultadoComparacion}
        <br><br>🔁 Puedes volver a lanzar la copia para intentar completar o actualizar los elementos.
        <br><br>🧹 Si deseas eliminar del destino los elementos que ya no están en el origen, usa el botón de <strong>Limpiar</strong>.`;
      esCoincidenteLocal = false;
    }

    return {
      mensaje: mensajeFinal,
      esCoincidente: esCoincidenteLocal,
      idDestino: idCarpetaDestino,
      urlRegistro: urlRegistro,
      nombreDestino: nombreDestinoReal
    };

  } catch (e) {
    Logger.log(`Error en verificarIntegridad para "${nombreDestinoReal}": <span class="math-inline">\{e\.message\}\\n</span>{e.stack}`);
    return { 
        mensaje: "❌ Error al verificar integridad: " + e.message, 
        esCoincidente: false, 
        idDestino: idCarpetaDestino, // Puede que lo tengamos si el error fue después de obtenerlo
        urlRegistro: urlRegistro, // Puede que lo tengamos
        nombreDestino: nombreDestinoReal
    };
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

    if (!ultimoArchivo) {
      // Devolver estructura consistente incluso si no hay log
      return { 
        registroJson: { /* podrías inicializar con estructura base si es necesario */
            __errores: [], __pesados: [], __pendientes: [], 
            __progreso: { totalTareas: 0, tareasCompletadas: 0, erroresEnLote: 0, mensajeActual: "No se encontró registro previo."}
        }, 
        archivoLog: null 
      };
    }
    const contenido = ultimoArchivo.getBlob().getDataAsString();
    return { registroJson: JSON.parse(contenido), archivoLog: ultimoArchivo };
  } catch (e) {
    Logger.log(`Error al cargar el registro para "${nombreDestino}": <span class="math-inline">\{e\.message\}\\n</span>{e.stack}`);
    // Devolver estructura consistente en caso de error
    return { 
      registroJson: {
            __errores: [`Error cargando registro: ${e.message}`], __pesados: [], __pendientes: [], 
            __progreso: { totalTareas: 0, tareasCompletadas: 0, erroresEnLote: 0, mensajeActual: "Error al cargar registro."}
      }, 
      archivoLog: null 
    };
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

