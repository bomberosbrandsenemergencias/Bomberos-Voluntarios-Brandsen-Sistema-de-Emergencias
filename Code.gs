/**
 * Sistema de Gestión de Emergencias — Bomberos Voluntarios Brandsen
 * Backend en Google Apps Script, atado a la planilla de Google Sheets que hace de base de datos.
 *
 * INSTALACIÓN (una sola vez):
 *
 * OPCIÓN A — el script vive DENTRO de la planilla (recomendado, más simple):
 * 1. Abrí (o creá) la planilla nueva en Google Sheets. Extensiones > Apps Script.
 * 2. Borrá lo que haya en Code.gs y pegá todo este archivo.
 * 3. Arriba, en el desplegable de funciones, elegí "inicializarPlanilla" y apretá Ejecutar.
 *    La primera vez va a pedir autorización (usa tu propia cuenta de Google, es normal el aviso de
 *    "app no verificada" — es tu propio script). Esto crea todas las pestañas con sus encabezados y
 *    un primer token de administrador en la pestaña "Accesos".
 *
 * OPCIÓN B — si ya creaste el proyecto de Apps Script suelto (desde script.google.com, sin abrirlo
 * desde una planilla), como este link que tenés ahora: hace falta decirle a qué planilla apunta.
 * 1. Pegá este archivo en Code.gs igual.
 * 2. Abrí (o creá) la planilla en Google Sheets, copiá el ID que aparece en su URL, entre
 *    "/d/" y "/edit" (una tira larga de letras y números).
 * 3. En el desplegable de funciones elegí "configurarPlanilla", apretá el ícono de relojito/engranaje
 *    junto a Ejecutar para pasarle un parámetro... en realidad más simple: abajo de todo en este
 *    archivo hay una función "configurarPlanillaAca()" — reemplazá adentro el texto
 *    "PEGÁ_ACÁ_EL_ID_DE_TU_PLANILLA" por ese ID, guardá, elegí "configurarPlanillaAca" en el
 *    desplegable y ejecutala una vez.
 * 4. Después ejecutá "inicializarPlanilla" una vez (mismo resultado que la Opción A).
 *
 * EN CUALQUIER CASO, PARA PUBLICARLO:
 * Implementar > Nueva implementación > tipo "Aplicación web".
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 *    Copiá la URL que te da (termina en /exec): esa es la dirección del backend que va a usar el sitio.
 * Cada vez que cambies este código tenés que crear una implementación NUEVA (Gestionar
 * implementaciones > lápiz > Versión: Nueva versión > Implementar) para que el cambio se vea
 * reflejado en esa misma URL.
 *
 * El sitio (frontend) le habla a este script por POST, con el cuerpo en JSON, pero con
 * Content-Type "text/plain" (no "application/json") para evitar que el navegador dispare un
 * preflight CORS que Apps Script no sabe responder. Esto ya está resuelto del lado del frontend,
 * no hace falta tocar nada acá por eso.
 */

// ---------- definición de colecciones ----------

var COLECCIONES = {
  Unidades: { columnas: ["id","numero","grupo"], json: [] },
  Personal: { columnas: ["id","nombre","dni","cargo","localidad","telefono","jerarquico","chofer","reserva","presente"], json: [] },
  Servicios: {
    columnas: ["id","fecha","horaLlamado","horaDespacho","horaRetorno","tipo","evento","solicitante",
      "direccion","lat","lng","localidad","barrio","unidades","personalPorUnidad","cantidadPersonal",
      "mayores","menores","evacuadosMayoresDetalle","evacuadosMenoresDetalle","centroEvacuacion",
      "observaciones","activo","retiradoDelCentro","creadoEn"],
    json: ["unidades","personalPorUnidad","evacuadosMayoresDetalle","evacuadosMenoresDetalle"]
  },
  CentrosEvacuacion: { columnas: ["id","nombre","estatus","direccion","responsable","telefono","capacidad","grupoElectrogeno","aguaPotable"], json: [] },
  IngresosDirectosCentro: {
    columnas: ["id","centro","tipoIngreso","evento","fecha","hora","mayores","menores",
      "evacuadosMayoresDetalle","evacuadosMenoresDetalle","observaciones","retiradoDelCentro","creadoEn"],
    json: ["evacuadosMayoresDetalle","evacuadosMenoresDetalle"]
  },
  RioLecturas: { columnas: ["id","distanciaMedida","lluviaAcumulada","observacion","fecha","hora","creadoEn"], json: [] },
  RiesgoLocalidades: { columnas: ["id","localidad","nivel","motivo"], json: [] },
  Helipuertos: { columnas: ["id","nombre","ubicacion","responsable","telefono","estado"], json: [] },
  Contactos: { columnas: ["id","organismo","nombre","cargo","telefono","notas"], json: [] },
  BitacoraEventos: { columnas: ["id","evento","texto","fecha","hora","creadoEn"], json: [] }
};

// pestañas "documento único" (una sola fila de datos, sin id)
var DOCUMENTOS = {
  Monitoreo: ["texto","fuente","actualizadoEn","umbralAlerta","umbralAlarma","lugarMedicion","distanciaNormalCauce","inicioLluviasFecha","inicioLluviasHora"],
  Catalogos: ["tipos","localidades","barrios"]
};
var DOCUMENTOS_JSON = { Catalogos: ["tipos","localidades","barrios"] };

var HOJA_ACCESOS = "Accesos";
var COLUMNAS_ACCESOS = ["token","rol","alcance","nombre","activo"];

// colecciones a las que el rol "centro" puede acceder, y cómo se filtra cada una por su alcance
var ALCANCE_ROL_CENTRO = {
  IngresosDirectosCentro: { campo: "centro", escritura: true },
  CentrosEvacuacion: { campo: "nombre", escritura: false },
  Servicios: { campo: "centroEvacuacion", escritura: false }
};

// para el rol "centro", qué columnas de Servicios puede ver (nada de personal/unidades cargadas)
var PROYECCION_ROL_CENTRO = {
  Servicios: ["id","fecha","hora","direccion","evento","mayores","menores",
    "evacuadosMayoresDetalle","evacuadosMenoresDetalle","centroEvacuacion","retiradoDelCentro"]
};

// valor por defecto de un campo JSON cuando está vacío, según su forma
var DEFAULT_OBJETO = { personalPorUnidad: {} };

// ---------- conexión con la planilla ----------

// funciona tanto si el script vive dentro de la planilla (Opción A) como si es un proyecto
// suelto al que se le indicó el ID de la planilla con configurarPlanilla() (Opción B)
//
// OJO RENDIMIENTO: abrir la planilla (SpreadsheetApp.openById/getActiveSpreadsheet) es, con
// diferencia, la llamada más lenta de todo este script. "listarTodo" la necesitaba una vez por
// cada una de las 10 colecciones + 2 documentos (13 aperturas de la misma planilla en una sola
// consulta, que además se repite sola cada 6 segundos desde cada celular/computadora conectado).
// Se cachea acá para abrirla una sola vez por ejecución (la referencia sigue "viva": leer/escribir
// a través de ella sigue trayendo datos al instante, no es una foto vieja).
var _planillaCache_ = null;
function obtenerPlanilla_() {
  if (_planillaCache_) return _planillaCache_;
  var id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (id) { _planillaCache_ = SpreadsheetApp.openById(id); return _planillaCache_; }
  var activa = SpreadsheetApp.getActiveSpreadsheet();
  if (activa) { _planillaCache_ = activa; return _planillaCache_; }
  throw new Error("No se encontró la planilla. Ejecutá configurarPlanillaAca() con el ID de tu Google Sheet (ver instrucciones arriba de todo).");
}

// mismo criterio para las pestañas individuales: dentro de una misma consulta, "listarTodo"
// pedía cada pestaña una sola vez así que esto no ahorra aperturas repetidas de LA MISMA hoja,
// pero deja un solo lugar para pedir cualquier hoja por nombre de forma consistente.
var _hojaCache_ = {};
function obtenerHoja_(nombre) {
  if (_hojaCache_[nombre]) return _hojaCache_[nombre];
  var hoja = obtenerPlanilla_().getSheetByName(nombre);
  _hojaCache_[nombre] = hoja;
  return hoja;
}

function configurarPlanilla(idPlanilla) {
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", idPlanilla);
}

// completá el ID entre comillas (lo sacás de la URL de tu planilla, entre "/d/" y "/edit") y
// ejecutá esta función una sola vez si tu script NO vive dentro de la planilla
function configurarPlanillaAca() {
  configurarPlanilla("PEGÁ_ACÁ_EL_ID_DE_TU_PLANILLA");
  Logger.log("Listo, planilla configurada. Ahora ejecutá inicializarPlanilla().");
}

// ---------- inicialización ----------

function inicializarPlanilla() {
  var ss = obtenerPlanilla_();
  Object.keys(COLECCIONES).forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(COLECCIONES[nombre].columnas);
      hoja.setFrozenRows(1);
    }
  });
  Object.keys(DOCUMENTOS).forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(DOCUMENTOS[nombre]);
      hoja.appendRow(new Array(DOCUMENTOS[nombre].length).fill(""));
      hoja.setFrozenRows(1);
    }
  });
  var hojaAcc = ss.getSheetByName(HOJA_ACCESOS) || ss.insertSheet(HOJA_ACCESOS);
  if (hojaAcc.getLastRow() === 0) {
    hojaAcc.appendRow(COLUMNAS_ACCESOS);
    hojaAcc.appendRow([Utilities.getUuid(), "admin", "", "Acceso general (renombrar)", true]);
    hojaAcc.setFrozenRows(1);
  }
  var porDefecto = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (porDefecto && porDefecto.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(porDefecto);
  }
  var mensaje = "Listo. Se crearon todas las pestañas.\n\n" +
    "Andá a la pestaña 'Accesos': ahí está el primer token (columna A) — es el que vas a usar vos " +
    "como administrador. Para dar acceso a un centro de evacuación, agregá una fila con rol 'centro' " +
    "y en 'alcance' el nombre EXACTO del centro.";
  try {
    SpreadsheetApp.getUi().alert(mensaje); // funciona si el script vive dentro de la planilla
  } catch (e) {
    Logger.log(mensaje); // proyecto suelto: mirá el resultado en "Ejecuciones" o con Ver > Registros
  }
}

// ---------- entrada HTTP ----------

function doGet(e) {
  return responder_({ ok: true, servicio: "Sistema de Gestión de Emergencias Brandsen", hora: new Date().toISOString() });
}

// el candado (LockService) solo hace falta para las escrituras, para que dos ediciones
// simultáneas no se pisen — usarlo también en las lecturas (como se hacía al principio)
// hace que el sondeo periódico del sitio (una lectura cada pocos segundos) se bloquee a sí
// mismo y tire "El sistema está ocupado" en cadena, sin que haya ninguna escritura real de por medio
var ACCIONES_ESCRITURA = { guardar: true, actualizar: true, actualizarLote: true, eliminar: true, guardarCatalogos: true, guardarMonitoreo: true };

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return responder_({ ok: false, error: "Cuerpo de la solicitud inválido." });
  }

  if (!ACCIONES_ESCRITURA[body.accion]) {
    // lectura: sin candado, para que varias consultas puedan resolverse en paralelo
    try {
      return responder_(procesar_(body));
    } catch (err) {
      return responder_({ ok: false, error: String(err) });
    }
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    // "ocupado: true" le permite al frontend reintentar esto solo unas pocas veces antes de
    // mostrárselo a quien está cargando datos, en vez de cortarlo en seco (ver comentario en
    // la función llamar() del frontend) -- pensado para el pico de varias personas cargando
    // servicios/ingresos a la vez durante una emergencia grande.
    return responder_({ ok: false, error: "El sistema está ocupado, probá de nuevo en un segundo.", ocupado: true });
  }
  try {
    var resultado = procesar_(body);
    // la escritura ya quedó guardada en la planilla: si había una respuesta de "listarTodo"
    // cacheada de HACE UN INSTANTE para este mismo token (ver acción listarTodo más abajo), hay
    // que actualizarla. Si no, el listarTodo que el propio frontend dispara justo después de
    // guardar (para refrescarse con lo recién guardado) podía traer esa versión vieja cacheada,
    // de ANTES de este cambio -- y lo que se acababa de tildar (ej. "Presente") se veía
    // destildar solo un instante después, aunque el guardado hubiera salido bien.
    if (resultado && resultado.ok) invalidarCacheListarTodo_(body);
    return responder_(resultado);
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Antes esto tiraba TODA la respuesta cacheada de "listarTodo" para ese rol+alcance apenas
// se guardaba cualquier cosa -- así, con varias personas cargando avisos seguidos (el caso de
// una emergencia grande), cada guardado obligaba a la SIGUIENTE consulta de CUALQUIERA que
// comparta ese rol+alcance (ej. otro admin) a releer las 10 hojas de la planilla enteras, aunque
// solo hubiera cambiado una. Ahora, en vez de tirar todo, se relee (barato: una sola hoja) nada
// más que la colección que efectivamente cambió y se actualiza esa parte dentro de la respuesta
// ya cacheada -- el resto sigue sirviéndose de la caché tal cual estaba.
function invalidarCacheListarTodo_(body) {
  try {
    var acceso = validarToken_(body.token);
    if (!acceso.ok) return;
    var cache = CacheService.getScriptCache();
    var cacheKey = "listarTodo_" + acceso.rol + "_" + (acceso.alcance || "");
    var cacheado = cache.get(cacheKey);
    if (!cacheado) return; // no había nada cacheado para este rol+alcance, no hay nada que parchar

    var datos = JSON.parse(cacheado);
    var coleccionCambiada = body.coleccion;
    if (coleccionCambiada && COLECCIONES[coleccionCambiada]) {
      var permisoCol = permisoColeccion_(acceso, coleccionCambiada, "listar");
      if (!permisoCol.ok) { cache.remove(cacheKey); return; }
      var filasCol = leerFilas_(obtenerHoja_(coleccionCambiada), COLECCIONES[coleccionCambiada]);
      if (permisoCol.filtroCampo) filasCol = filasCol.filter(function(f) { return f[permisoCol.filtroCampo] === acceso.alcance; });
      if (permisoCol.proyeccion) {
        filasCol = filasCol.map(function(f) {
          var out = {};
          permisoCol.proyeccion.forEach(function(c) { out[c] = f[c]; });
          return out;
        });
      }
      datos[coleccionCambiada] = filasCol;
    } else if (body.accion === "guardarMonitoreo" && acceso.rol === "admin") {
      datos.Monitoreo = leerDocumentoRaw_("Monitoreo");
    } else if (body.accion === "guardarCatalogos" && acceso.rol === "admin") {
      datos.Catalogos = leerDocumentoRaw_("Catalogos");
    } else {
      // acción no reconocida acá (no debería pasar): por las dudas, mejor no dejar algo
      // desactualizado cacheado que quedarse con una parte sin poder identificar qué cambió
      cache.remove(cacheKey);
      return;
    }
    try { cache.put(cacheKey, JSON.stringify(datos), 5); } catch (e) { cache.remove(cacheKey); }
  } catch (e) {
    // si esto llegara a fallar, en el peor caso el próximo sondeo tarda hasta 5s de más en
    // reflejar el cambio (se soluciona solo cuando la caché vence) -- no rompe nada más.
  }
}

function responder_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- lógica principal ----------

function procesar_(body) {
  var token = body.token || "";
  var accion = body.accion;

  var acceso = validarToken_(token);
  if (!acceso.ok) return { ok: false, error: acceso.error || "Acceso inválido." };

  if (accion === "sesion") return { ok: true, rol: acceso.rol, alcance: acceso.alcance, nombre: acceso.nombre };

  if (accion === "leerCatalogos" || accion === "guardarCatalogos") {
    return procesarDocumento_("Catalogos", accion === "guardarCatalogos" ? "guardar" : "leer", body, acceso);
  }
  if (accion === "leerMonitoreo" || accion === "guardarMonitoreo") {
    return procesarDocumento_("Monitoreo", accion === "guardarMonitoreo" ? "guardar" : "leer", body, acceso);
  }

  // una sola llamada que trae todas las colecciones (y, para admin, los documentos) a las que
  // el rol tiene acceso — para no hacer una llamada por colección al sondear cambios (polling)
  if (accion === "listarTodo") {
    // OJO RENDIMIENTO (2): con varios dispositivos conectados a la vez (ej. un admin + dos o tres
    // centros de evacuación), cada uno sondea cada 6s y hasta ahora cada sondeo releía las 10
    // hojas de nuevo aunque los datos no hubieran cambiado un solo segundo antes. Se cachea acá
    // la respuesta completa por 5 segundos (menos que el intervalo de sondeo de 6s), separada por
    // rol+alcance (un admin nunca comparte caché con un centro, ni un centro con otro): así, si dos
    // sondeos caen dentro de esa ventana de 5s, el segundo se resuelve al instante sin volver a leer
    // la planilla. Si la respuesta es demasiado grande para cachear (>100KB, poco probable), se seguía
    // funcionando igual, solo que sin este ahorro puntual.
    var cacheKey = "listarTodo_" + acceso.rol + "_" + (acceso.alcance || "");
    var cache = CacheService.getScriptCache();
    var cacheado = cache.get(cacheKey);
    if (cacheado) return { ok: true, datos: JSON.parse(cacheado) };

    var resultado = {};
    Object.keys(COLECCIONES).forEach(function(nombreCol) {
      var permisoCol = permisoColeccion_(acceso, nombreCol, "listar");
      if (!permisoCol.ok) return; // esta colección no es visible para el rol, se omite
      var hojaCol = obtenerHoja_(nombreCol);
      var defCol = COLECCIONES[nombreCol];
      var filasCol = leerFilas_(hojaCol, defCol);
      if (permisoCol.filtroCampo) {
        filasCol = filasCol.filter(function(f) { return f[permisoCol.filtroCampo] === acceso.alcance; });
      }
      if (permisoCol.proyeccion) {
        filasCol = filasCol.map(function(f) {
          var out = {};
          permisoCol.proyeccion.forEach(function(c) { out[c] = f[c]; });
          return out;
        });
      }
      resultado[nombreCol] = filasCol;
    });
    if (acceso.rol === "admin") {
      resultado.Monitoreo = leerDocumentoRaw_("Monitoreo");
      resultado.Catalogos = leerDocumentoRaw_("Catalogos");
    }
    try { cache.put(cacheKey, JSON.stringify(resultado), 5); } catch (e) { /* respuesta muy grande para cachear, seguimos sin caché */ }
    return { ok: true, datos: resultado };
  }

  // actualización en bloque: varios documentos de UNA colección en una sola llamada al backend,
  // en vez de una llamada HTTP por documento (que es lo que hacían operaciones como "Marcar
  // todos no presente" o "Aplicar evento a los seleccionados" — con 20-30 personas/servicios
  // eso significaba 20-30 idas y vueltas seguidas al servidor, cada una con el costo fijo de
  // abrir la planilla y esperar el candado, y por eso se sentían MUY lentas). Acá se abre la
  // hoja una sola vez, se lee y se reescribe en una sola pasada.
  if (accion === "actualizarLote") {
    var coleccionLote = body.coleccion;
    if (!coleccionLote || !COLECCIONES[coleccionLote]) return { ok: false, error: "Colección desconocida." };
    var permisoLote = permisoColeccion_(acceso, coleccionLote, "actualizar");
    if (!permisoLote.ok) return { ok: false, error: permisoLote.error || "No autorizado." };
    var resultadoLote = actualizarFilasLote_(obtenerHoja_(coleccionLote), COLECCIONES[coleccionLote], body.cambios || [], permisoLote, acceso);
    return { ok: true, actualizados: resultadoLote.actualizados, noEncontrados: resultadoLote.noEncontrados };
  }

  var coleccion = body.coleccion;
  if (!coleccion || !COLECCIONES[coleccion]) return { ok: false, error: "Colección desconocida." };

  var permiso = permisoColeccion_(acceso, coleccion, accion);
  if (!permiso.ok) return { ok: false, error: permiso.error || "No autorizado." };

  var hoja = obtenerHoja_(coleccion);
  var def = COLECCIONES[coleccion];

  if (accion === "listar") {
    var filas = leerFilas_(hoja, def);
    if (permiso.filtroCampo) {
      filas = filas.filter(function(f) { return f[permiso.filtroCampo] === acceso.alcance; });
    }
    if (permiso.proyeccion) {
      filas = filas.map(function(f) {
        var out = {};
        permiso.proyeccion.forEach(function(c) { out[c] = f[c]; });
        return out;
      });
    }
    return { ok: true, datos: filas };
  }

  if (accion === "guardar") { // upsert: crea si el id no existe, reemplaza si existe
    var datosNuevo = body.datos || {};
    if (permiso.filtroCampo) datosNuevo[permiso.filtroCampo] = acceso.alcance;
    // "nuevo: true" lo manda el frontend SOLO cuando el id se acaba de generar acá mismo (ver
    // collection().add() en index.html) -- en ese caso no puede existir todavía, así que nos
    // ahorramos el escaneo de la columna de ids para buscarlo (buscarFila_ adentro de
    // guardarFila_) y vamos directo a agregar la fila. Es la escritura más frecuente durante una
    // emergencia grande (varios "Nuevo aviso" seguidos), así que achicar el trabajo bajo el
    // candado ahí ayuda más que en el resto.
    guardarFila_(hoja, def, body.id, datosNuevo, !!body.nuevo);
    return { ok: true };
  }

  if (accion === "actualizar") { // merge parcial sobre una fila existente
    var filaIdx = buscarFila_(hoja, body.id);
    if (filaIdx === -1) return { ok: false, error: "No existe ese registro." };
    if (permiso.filtroCampo) {
      var actual = leerFilaPorIndice_(hoja, def, filaIdx);
      if (actual[permiso.filtroCampo] !== acceso.alcance) return { ok: false, error: "No autorizado." };
    }
    actualizarFila_(hoja, def, filaIdx, body.datos || {});
    return { ok: true };
  }

  if (accion === "eliminar") {
    var filaIdx2 = buscarFila_(hoja, body.id);
    if (filaIdx2 === -1) return { ok: true }; // ya no está, lo damos por hecho
    if (permiso.filtroCampo) {
      var actual2 = leerFilaPorIndice_(hoja, def, filaIdx2);
      if (actual2[permiso.filtroCampo] !== acceso.alcance) return { ok: false, error: "No autorizado." };
    }
    hoja.deleteRow(filaIdx2);
    return { ok: true };
  }

  return { ok: false, error: "Acción desconocida." };
}

function procesarDocumento_(nombre, tipo, body, acceso) {
  if (acceso.rol !== "admin") return { ok: false, error: "No autorizado." };

  if (tipo === "leer") {
    return { ok: true, datos: leerDocumentoRaw_(nombre) };
  }

  // guardar: merge de lo nuevo sobre lo que ya había
  var hoja = obtenerHoja_(nombre);
  var columnas = DOCUMENTOS[nombre];
  var jsonCols = DOCUMENTOS_JSON[nombre] || [];
  var actual = leerDocumentoRaw_(nombre);
  var nuevo = Object.assign({}, actual, body.datos || {});
  var fila = columnas.map(function(c) {
    var v = nuevo[c];
    return jsonCols.indexOf(c) > -1 ? JSON.stringify(v || []) : v;
  });
  hoja.getRange(2, 1, 1, columnas.length).setValues([fila]);
  return { ok: true };
}

// lee una pestaña "documento único" (una sola fila de datos) y la devuelve como objeto plano,
// parseando las columnas que son JSON. La usan tanto procesarDocumento_ (leerMonitoreo/leerCatalogos)
// como listarTodo (para incluir estos documentos junto con las colecciones en una sola respuesta).
function leerDocumentoRaw_(nombre) {
  var hoja = obtenerHoja_(nombre);
  var columnas = DOCUMENTOS[nombre];
  var jsonCols = DOCUMENTOS_JSON[nombre] || [];
  var valores = hoja.getRange(2, 1, 1, columnas.length).getValues()[0];
  var obj = {};
  columnas.forEach(function(c, i) {
    var v = normalizarValorCelda_(c, valores[i]);
    obj[c] = jsonCols.indexOf(c) > -1 ? parsearJSON_(v) : v;
  });
  return obj;
}

// Google Sheets puede llegar a guardar un texto simple como "14:23" como un valor interno de
// Hora/Fecha, si en algún momento detectó ese "formato" en la columna (basta con que una sola
// celda se haya cargado o pegado alguna vez de forma que Sheets la interpretara así). Cuando eso
// pasa, Apps Script devuelve esa celda como un objeto Date de JavaScript (con el 30/31 de
// diciembre de 1899 como "fecha" de referencia para un valor que en realidad es solo una hora),
// y si ese objeto se manda tal cual al frontend termina viéndose como "1899-12-31T02:01:48.000Z"
// en vez de la hora que realmente se cargó (esto es justo lo que reportó Fede). Acá se lo vuelve
// a pasar a texto plano, según qué tipo de dato es esa columna por su nombre.
function normalizarValorCelda_(nombreColumna, valor) {
  if (!(valor instanceof Date)) return valor;
  var tz = Session.getScriptTimeZone();
  var nombreMin = String(nombreColumna).toLowerCase();
  if (nombreMin.indexOf("hora") > -1) return Utilities.formatDate(valor, tz, "HH:mm");
  if (nombreMin.indexOf("fecha") > -1) return Utilities.formatDate(valor, tz, "yyyy-MM-dd");
  return valor.toISOString();
}

// ---------- autorización ----------

function validarToken_(token) {
  if (!token) return { ok: false, error: "Falta el token de acceso." };
  var hoja = obtenerHoja_(HOJA_ACCESOS);
  var valores = hoja.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === token) {
      var activo = valores[i][4];
      if (activo !== true && String(activo).toUpperCase() !== "TRUE") {
        return { ok: false, error: "Este acceso fue desactivado." };
      }
      return { ok: true, rol: valores[i][1], alcance: valores[i][2], nombre: valores[i][3] };
    }
  }
  return { ok: false, error: "Token inválido." };
}

function permisoColeccion_(acceso, coleccion, accion) {
  if (acceso.rol === "admin") return { ok: true };
  if (acceso.rol === "centro") {
    var regla = ALCANCE_ROL_CENTRO[coleccion];
    if (!regla) return { ok: false, error: "Sin acceso a esta sección." };
    var esEscritura = accion === "guardar" || accion === "actualizar" || accion === "eliminar";
    if (esEscritura && !regla.escritura) return { ok: false, error: "Solo lectura para tu acceso." };
    return { ok: true, filtroCampo: regla.campo, proyeccion: PROYECCION_ROL_CENTRO[coleccion] || null };
  }
  return { ok: false, error: "Rol desconocido." };
}

// ---------- lectura/escritura genérica de filas ----------

function leerFilas_(hoja, def) {
  var valores = hoja.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < valores.length; i++) {
    if (!valores[i][0]) continue;
    out.push(filaAObjeto_(valores[i], def));
  }
  return out;
}

function leerFilaPorIndice_(hoja, def, filaIdx) {
  var valores = hoja.getRange(filaIdx, 1, 1, def.columnas.length).getValues()[0];
  return filaAObjeto_(valores, def);
}

function filaAObjeto_(valores, def) {
  var obj = {};
  def.columnas.forEach(function(c, i) {
    var v = normalizarValorCelda_(c, valores[i]);
    obj[c] = def.json.indexOf(c) > -1 ? parsearJSON_(v) : v;
  });
  // "id" siempre viaja como texto: si el id es numérico (ej. un DNI), Sheets lo guarda/lee
  // como número, y el frontend compara ids con === contra el string que puso en el HTML
  // (data-presente="..." y similares) — un número nunca es === a ese string, así que el click
  // quedaba sin efecto en silencio. Firestore (el sistema original) siempre daba ids como texto.
  if (obj.id !== undefined && obj.id !== null && obj.id !== "") obj.id = String(obj.id);
  return obj;
}

function parsearJSON_(valor) {
  if (valor === "" || valor === null || valor === undefined) return [];
  if (typeof valor !== "string") return valor;
  try { return JSON.parse(valor); } catch (e) { return []; }
}

function buscarFila_(hoja, id) {
  var n = Math.max(hoja.getLastRow() - 1, 0);
  if (n === 0) return -1;
  var ids = hoja.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function valorJSONPorDefecto_(campo) {
  return DEFAULT_OBJETO.hasOwnProperty(campo) ? DEFAULT_OBJETO[campo] : [];
}

function guardarFila_(hoja, def, id, datos, esNuevo) {
  datos = datos || {};
  datos.id = id;
  var fila = def.columnas.map(function(c) {
    var v = datos[c];
    if (def.json.indexOf(c) > -1) return JSON.stringify(v === undefined ? valorJSONPorDefecto_(c) : v);
    return v === undefined ? "" : v;
  });
  if (esNuevo) { hoja.appendRow(fila); return; } // ver comentario en el llamador (acción "guardar")
  var filaIdx = buscarFila_(hoja, id);
  if (filaIdx === -1) {
    hoja.appendRow(fila);
  } else {
    hoja.getRange(filaIdx, 1, 1, fila.length).setValues([fila]);
  }
}

// versión en bloque de actualizarFila_: recibe una lista de {id, datos} para la MISMA hoja y
// aplica todos los cambios con una sola lectura y una sola escritura del rango completo, en vez
// de una lectura+escritura por cada id (ver comentario en procesar_, acción "actualizarLote").
function actualizarFilasLote_(hoja, def, cambios, permiso, acceso) {
  if (!cambios || !cambios.length) return { actualizados: 0, noEncontrados: [] };
  var n = Math.max(hoja.getLastRow() - 1, 0);
  if (n === 0) return { actualizados: 0, noEncontrados: cambios.map(function(c) { return c.id; }) };

  var rango = hoja.getRange(2, 1, n, def.columnas.length);
  var valores = rango.getValues();
  var indicePorId = {};
  valores.forEach(function(fila, i) { indicePorId[String(fila[0])] = i; });

  var noEncontrados = [];
  var huboCambios = false;
  cambios.forEach(function(cambio) {
    var i = indicePorId[String(cambio.id)];
    if (i === undefined) { noEncontrados.push(cambio.id); return; }
    var actual = filaAObjeto_(valores[i], def);
    if (permiso.filtroCampo && actual[permiso.filtroCampo] !== acceso.alcance) { noEncontrados.push(cambio.id); return; }
    var nuevo = Object.assign({}, actual, cambio.datos || {});
    valores[i] = def.columnas.map(function(c) {
      var v = nuevo[c];
      if (def.json.indexOf(c) > -1) return JSON.stringify(v === undefined ? valorJSONPorDefecto_(c) : v);
      return v === undefined ? "" : v;
    });
    huboCambios = true;
  });
  if (huboCambios) rango.setValues(valores);
  return { actualizados: cambios.length - noEncontrados.length, noEncontrados: noEncontrados };
}

function actualizarFila_(hoja, def, filaIdx, cambios) {
  var actual = leerFilaPorIndice_(hoja, def, filaIdx);
  var nuevo = Object.assign({}, actual, cambios);
  var fila = def.columnas.map(function(c) {
    var v = nuevo[c];
    if (def.json.indexOf(c) > -1) return JSON.stringify(v === undefined ? valorJSONPorDefecto_(c) : v);
    return v === undefined ? "" : v;
  });
  hoja.getRange(filaIdx, 1, 1, fila.length).setValues([fila]);
}
