// =========================================================================
// CONFIGURACIÓN GLOBAL
// =========================================================================
const SPREADSHEET_ID  = '1XpEYx1c10W1itU12pRsmtAF1Skb6otsVMGlmVYNJqms';
const CARPETA_RAIZ_ID = '1gH9Yj-zE6hPK5Avn2JwOYSjh39aOJKth';
const PREFIJO_FOLIO   = 'AD';

// =========================================================================
// PUNTO DE ENTRADA WEB
// =========================================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('REGISTRO ADMISIÓN PAAE 2026-2027')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =========================================================================
// GENERACIÓN DE FOLIO ÚNICO ANTI-CONCURRENCIA
// =========================================================================
function generarFolio() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    let hoja  = ss.getSheetByName('Control_Folios');
    if (!hoja) {
      hoja = ss.insertSheet('Control_Folios');
      hoja.getRange('A1').setValue('ULTIMO_FOLIO');
      hoja.getRange('B1').setValue(0);
    }
    const ultimo = Number(hoja.getRange('B1').getValue()) || 0;
    const nuevo  = ultimo + 1;
    hoja.getRange('B1').setValue(nuevo);
    SpreadsheetApp.flush();
    return PREFIJO_FOLIO + String(nuevo).padStart(4, '0');
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// FUNCIÓN PRINCIPAL DE REGISTRO
// =========================================================================
function procesarRegistro(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaAspirantes    = ss.getSheetByName('Aspirantes_PAAE');
    const hojaPeriodos      = ss.getSheetByName('Periodos_PAAE');
    const hojaActualizacion = ss.getSheetByName('Actualizacion_PAAE');

    if (!hojaAspirantes || !hojaPeriodos || !hojaActualizacion) {
      return { exito: false, mensaje: 'Error de configuración: una o más hojas no encontradas.' };
    }

    const curp = formData.datosIdentificacion.curp.trim().toUpperCase();

    // ── Validar duplicado de CURP ──────────────────────────────────────────
    if (existeCURP(hojaAspirantes, curp)) {
      return {
        exito: false,
        mensaje: 'Ya existe un registro con esta CURP. Cualquier corrección se realizará en su cita de validación.'
      };
    }

    // ── Validar RFC ────────────────────────────────────────────────────────
    const rfc = formData.datosIdentificacion.rfc.trim().toUpperCase();
    if (rfc.length !== 13) {
      return { exito: false, mensaje: 'El RFC de persona física debe tener exactamente 13 caracteres.' };
    }

    // ── Generar folio único ────────────────────────────────────────────────
    const folio = generarFolio();

    // ── Normalizar fechas ──────────────────────────────────────────────────
    formData.periodos = formData.periodos.map(function(p) {
      return {
        fechaInicio:   normalizarFecha(p.fechaInicio),
        fechaFin:      normalizarFecha(p.fechaFin),
        funcion:       p.funcion,
        tipoDocumento: p.tipoDocumento || ''
      };
    });
    formData.actualizacion = formData.actualizacion.map(function(c) {
      return {
        tipo:        c.tipo,
        nombre:      c.nombre,
        institucion: c.institucion,
        fecha:       normalizarFecha(c.fecha)
      };
    });

    // ── Crear carpetas en Drive ────────────────────────────────────────────
    const carpetaAspirante = crearCarpetasAspirante(folio, curp, formData.datosPersonales.nombreInstitucional);

    // ── Guardar PDFs ───────────────────────────────────────────────────────
    const urlsArchivos = guardarArchivos(carpetaAspirante, formData, curp, folio);

    // ── Cálculos ───────────────────────────────────────────────────────────
    const calculoAntiguedad = calcularAntiguedadYpuntajes(
      formData.periodos,
      formData.datosLaborales.funcion
    );
    const puntajes = calcularPuntajesMultifactoriales({
      antiguedad:    calculoAntiguedad,
      funcion:       formData.datosLaborales.funcion,
      formacion:     formData.formacion,
      actualizacion: formData.actualizacion,
      carta:         formData.carta
    });

    // ── Registrar en hojas ─────────────────────────────────────────────────
    registrarAspirante(hojaAspirantes, formData, calculoAntiguedad, puntajes, urlsArchivos, folio);
    registrarPeriodos(hojaPeriodos, curp, folio, formData.periodos, urlsArchivos.urlsPeriodos);
    if (formData.actualizacion.length > 0) {
      registrarActualizacion(hojaActualizacion, curp, folio, formData.actualizacion, urlsArchivos.urlsActualizacion);
    }

    // ── Correo de confirmación ─────────────────────────────────────────────
    enviarCorreoConfirmacion(formData, calculoAntiguedad, curp, folio, urlsArchivos);

    return {
      exito:   true,
      folio:   folio,
      mensaje: '✅ Expediente registrado correctamente. Su folio de registro es: ' + folio + '. Recibirá un correo de confirmación en breve.'
    };

  } catch (error) {
    Logger.log('Error en procesarRegistro: ' + error.toString());
    return { exito: false, mensaje: 'Error interno: ' + error.toString() };
  }
}

// =========================================================================
// NORMALIZACIÓN DE FECHA
// =========================================================================
function normalizarFecha(val) {
  if (!val) return '';
  val = val.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(val)) {
    const p = val.split('/');
    return p[2] + '-' + p[1] + '-' + p[0];
  }
  return val;
}

// =========================================================================
// VALIDACIÓN DE DUPLICADOS
// =========================================================================
function existeCURP(hoja, curp) {
  const datos = hoja.getDataRange().getValues();
  if (datos.length <= 1) return false;
  // Buscar columna CURP por encabezado (robusto ante reordenamiento)
  const idxCURP = datos[0].indexOf('CURP');
  if (idxCURP === -1) return false;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idxCURP]).trim().toUpperCase() === curp) return true;
  }
  return false;
}

// =========================================================================
// CREACIÓN DE CARPETAS EN DRIVE
// =========================================================================
function crearCarpetasAspirante(folio, curp, nombreCompleto) {
  const carpetaRaiz   = DriveApp.getFolderById(CARPETA_RAIZ_ID);
  const nombre        = folio + '_' + curp + '_' + nombreCompleto.replace(/ /g,'_').replace(/\//g,'-');
  const carpeta       = carpetaRaiz.createFolder(nombre);
  carpeta.createFolder('01_Identificacion');
  carpeta.createFolder('02_Formacion_Academica');
  carpeta.createFolder('03_Antiguedad_Soportes');
  carpeta.createFolder('04_Actualizacion_Desarrollo');
  carpeta.createFolder('05_Carta_Reconocimiento');
  return carpeta;
}

// =========================================================================
// GUARDADO DE ARCHIVOS PDF EN DRIVE
// Nombres incluyen folio, CURP y tipo de documento
// =========================================================================
function guardarArchivos(carpetaAspirante, formData, curp, folio) {
  const resultado = {
    identificacion:    '',
    titulo:            '',
    cedula:            '',
    certificado:       '',
    boleta:            '',
    cursoNinias:       '',
    certifNinias:      '',
    carta:             '',
    urlsPeriodos:      [],
    nombresPeriodos:   [],
    urlsActualizacion: [],
    nombresActualizacion: []
  };

  const map = {};
  const it  = carpetaAspirante.getFolders();
  while (it.hasNext()) { const f = it.next(); map[f.getName()] = f; }

  function pdf(carpeta, b64, nombre) {
    if (!b64) return '';
    try {
      const data = b64.indexOf(',') !== -1 ? b64.split(',')[1] : b64;
      const blob = Utilities.newBlob(Utilities.base64Decode(data), 'application/pdf', nombre);
      const arch = carpeta.createFile(blob);
      arch.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return arch.getUrl();
    } catch(e) { Logger.log('PDF error ' + nombre + ': ' + e); return ''; }
  }

  const pre = folio + '_' + curp;

  // 01 Identificación
  if (formData.archivosRaw.identificacion) {
    const nom = pre + '_IDENTIFICACION.pdf';
    resultado.identificacion = pdf(map['01_Identificacion'], formData.archivosRaw.identificacion, nom);
  }

  // 02 Formación Académica
  const fAc = map['02_Formacion_Academica'];
  if (formData.archivosRaw.titulo) {
    const nom = pre + '_TITULO_PROFESIONAL.pdf';
    resultado.titulo = pdf(fAc, formData.archivosRaw.titulo, nom);
  }
  if (formData.archivosRaw.cedula) {
    const nom = pre + '_CEDULA_PROFESIONAL.pdf';
    resultado.cedula = pdf(fAc, formData.archivosRaw.cedula, nom);
  }
  if (formData.archivosRaw.certificado) {
    const nom = pre + '_CERTIFICADO_ESTUDIOS.pdf';
    resultado.certificado = pdf(fAc, formData.archivosRaw.certificado, nom);
  }
  if (formData.archivosRaw.boleta) {
    const nom = pre + '_BOLETA_CALIFICACIONES.pdf';
    resultado.boleta = pdf(fAc, formData.archivosRaw.boleta, nom);
  }
  if (formData.archivosRaw.cursoNinias) {
    const nom = pre + '_CURSO_PRACTICO_AFIN.pdf';
    resultado.cursoNinias = pdf(fAc, formData.archivosRaw.cursoNinias, nom);
  }
  if (formData.archivosRaw.certifNinias) {
    const nom = pre + '_CERTIFICACION_OFICIAL_AFIN.pdf';
    resultado.certifNinias = pdf(fAc, formData.archivosRaw.certifNinias, nom);
  }

  // 03 Períodos — el nombre incluye el tipo de documento
  formData.periodos.forEach(function(p, i) {
    const tipoDoc = (p.tipoDocumento || 'DOCUMENTO').replace(/ /g,'_');
    const nom     = pre + '_PERIODO_' + (i + 1) + '_' + tipoDoc + '.pdf';
    resultado.nombresPeriodos.push(nom);
    const url = (formData.archivosRaw.periodosFiles && formData.archivosRaw.periodosFiles[i])
      ? pdf(map['03_Antiguedad_Soportes'], formData.archivosRaw.periodosFiles[i], nom)
      : '';
    resultado.urlsPeriodos.push(url);
  });

  // 04 Actualización
  formData.actualizacion.forEach(function(c, i) {
    const nom = pre + '_ACTUALIZACION_' + (i + 1) + '_' + (c.tipo || 'CONSTANCIA') + '.pdf';
    resultado.nombresActualizacion.push(nom);
    const url = (formData.archivosRaw.actualizacionFiles && formData.archivosRaw.actualizacionFiles[i])
      ? pdf(map['04_Actualizacion_Desarrollo'], formData.archivosRaw.actualizacionFiles[i], nom)
      : '';
    resultado.urlsActualizacion.push(url);
  });

  // 05 Carta de Reconocimiento
  if (formData.archivosRaw.carta) {
    const nom = pre + '_CARTA_RECONOCIMIENTO.pdf';
    resultado.carta = pdf(map['05_Carta_Reconocimiento'], formData.archivosRaw.carta, nom);
  }

  return resultado;
}

// =========================================================================
// CÁLCULO DE ANTIGÜEDAD
// =========================================================================
function calcularAntiguedadYpuntajes(periodos, funcionParticipa) {
  const rangoInicio = new Date('2021-01-01T00:00:00');
  const rangoFin    = new Date('2026-12-31T00:00:00');
  let totalMesesSEGEY = 0, totalMesesRango = 0, totalMesesFuncion = 0;

  periodos.forEach(function(p) {
    if (!p.fechaInicio || !p.fechaFin) return;
    const ini = new Date(p.fechaInicio + 'T00:00:00');
    const fin = new Date(p.fechaFin   + 'T00:00:00');
    if (isNaN(ini.getTime()) || isNaN(fin.getTime()) || ini > fin) return;
    const meses = diffMeses(ini, fin);
    totalMesesSEGEY += meses;
    const iniR = ini < rangoInicio ? rangoInicio : ini;
    const finR = fin > rangoFin   ? rangoFin   : fin;
    if (iniR <= finR) totalMesesRango += diffMeses(iniR, finR);
    if (p.funcion === funcionParticipa) totalMesesFuncion += meses;
  });

  function descomponer(meses) {
    const anios  = Math.floor(meses / 12);
    const mesesR = meses % 12;
    const puntaje = Math.min(30, anios * 3 + mesesR * 0.25);
    return { anios: anios, meses: mesesR, puntaje: puntaje };
  }

  return {
    segey:        descomponer(totalMesesSEGEY),
    rango:        descomponer(totalMesesRango),
    funcionEspec: descomponer(totalMesesFuncion)
  };
}

function diffMeses(ini, fin) {
  const y = fin.getFullYear() - ini.getFullYear();
  const m = fin.getMonth()   - ini.getMonth();
  const d = fin.getDate()    - ini.getDate();
  return y * 12 + m + (d >= 0 ? 0 : -1);
}

// =========================================================================
// CÁLCULO DE PUNTAJES
// =========================================================================
function calcularPuntajesMultifactoriales(datos) {
  const pAntiguedad     = datos.antiguedad.segey.puntaje;
  const pReconocimiento = Number(datos.carta.puntaje) || 0;
  let pFormacion = 0;
  const funcion = datos.funcion;
  const nivel   = datos.formacion.nivel || '';
  if (funcion === 'AUXILIAR DE INTENDENCIA') {
    if      (nivel.includes('TERCER') || nivel.includes('CUARTO') || nivel.includes('QUINTO')) pFormacion = 5;
    else if (nivel.includes('PRIMARIA'))    pFormacion = 10;
    else if (nivel.includes('SECUNDARIA'))  pFormacion = 15;
    else if (nivel.includes('BACHILLERATO') || nivel.includes('TÉCNICA')) pFormacion = 20;
    else if (nivel.includes('SUPERIOR UNIVERSITARIO'))                    pFormacion = 25;
  } else {
    if      (nivel.includes('BACHILLERATO') || nivel.includes('TÉCNICA')) pFormacion = 10;
    else if (nivel.includes('SUPERIOR UNIVERSITARIO'))                    pFormacion = 20;
    else if (nivel.includes('LICENCIATURA') || nivel.includes('INGENIERÍA')) pFormacion = 25;
  }
  let pActBase = 0, pActBono = 0;
  datos.actualizacion.forEach(function(c) {
    if      (c.tipo === 'TALLER')                                  pActBase += 1.5;
    else if (c.tipo === 'CURSO')                                   pActBase += 2.0;
    else if (c.tipo === 'DIPLOMADO' || c.tipo === 'CERTIFICACION') pActBase += 3.0;
    const instHigh = ['SEGEY','UPN','SEP','ICATEY','CECATI'];
    const instMed  = ['CONALEP','SLIM','MEXICOX','APRENDEMX'];
    if      (instHigh.includes(c.institucion)) pActBono += 0.5;
    else if (instMed.includes(c.institucion))  pActBono += 0.25;
  });
  const pActualizacion = Math.min(18, Math.min(15, pActBase) + Math.min(3, pActBono));
  return {
    antiguedad:     pAntiguedad,
    reconocimiento: pReconocimiento,
    formacion:      pFormacion,
    actualizacion:  pActualizacion,
    total:          pAntiguedad + pReconocimiento + pFormacion + pActualizacion
  };
}

// =========================================================================
// REGISTRO EN HOJAS DE CÁLCULO
// Detecta si ya existe encabezado con FOLIO para no duplicar
// =========================================================================
function registrarAspirante(hoja, formData, antiguedad, puntajes, urls, folio) {
  const encabezados = [
    'FOLIO','FECHA_REGISTRO','AP_PATERNO','AP_MATERNO','NOMBRES',
    'NOMBRE_INSTITUCIONAL','NOMBRE_NATURAL',
    'CURP','RFC','CORREO','TELEFONO',
    'FUNCION','NIVEL_ACADEMICO','PROMEDIO',
    'ANIOS_SEGEY','MESES_SEGEY','PTJE_ANTIGUEDAD',
    'ANIOS_RANGO','MESES_RANGO','PTJE_RANGO',
    'ANIOS_FUNCION','MESES_FUNCION','PTJE_FUNCION',
    'PTJE_RECONOCIMIENTO','PTJE_FORMACION','PTJE_ACTUALIZACION','PTJE_TOTAL',
    'URL_IDENTIFICACION','URL_CARTA'
  ];
  // Si la hoja está vacía o el primer encabezado no es FOLIO → escribir encabezados
  const primeraFila = hoja.getLastRow() > 0 ? hoja.getRange(1,1).getValue() : '';
  if (hoja.getLastRow() === 0 || primeraFila !== 'FOLIO') {
    if (hoja.getLastRow() > 0) hoja.insertRowBefore(1); // insertar fila antes de datos existentes
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }
  hoja.appendRow([
    folio, new Date(),
    formData.datosPersonales.apellidoPaterno,
    formData.datosPersonales.apellidoMaterno,
    formData.datosPersonales.nombres,
    formData.datosPersonales.nombreInstitucional,
    formData.datosPersonales.nombreNatural,
    formData.datosIdentificacion.curp,
    formData.datosIdentificacion.rfc,
    formData.datosIdentificacion.correo,
    formData.datosIdentificacion.telefono,
    formData.datosLaborales.funcion,
    formData.formacion.nivel,
    formData.formacion.promedio || '',   // vacío si no se ingresó
    antiguedad.segey.anios, antiguedad.segey.meses,  antiguedad.segey.puntaje,
    antiguedad.rango.anios, antiguedad.rango.meses,  antiguedad.rango.puntaje,
    antiguedad.funcionEspec.anios, antiguedad.funcionEspec.meses, antiguedad.funcionEspec.puntaje,
    puntajes.reconocimiento, puntajes.formacion, puntajes.actualizacion, puntajes.total,
    urls.identificacion, urls.carta
  ]);
}

function registrarPeriodos(hoja, curp, folio, periodos, urls) {
  const encabezados = ['FOLIO','CURP','NUM_PERIODO','FECHA_INICIO','FECHA_FIN','FUNCION','TIPO_DOCUMENTO','URL_SOPORTE'];
  const primeraFila = hoja.getLastRow() > 0 ? hoja.getRange(1,1).getValue() : '';
  if (hoja.getLastRow() === 0 || primeraFila !== 'FOLIO') {
    if (hoja.getLastRow() > 0) hoja.insertRowBefore(1);
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }
  periodos.forEach(function(p, i) {
    hoja.appendRow([folio, curp, i+1, p.fechaInicio, p.fechaFin, p.funcion, p.tipoDocumento || '', urls[i] || '']);
  });
}

function registrarActualizacion(hoja, curp, folio, actualizacion, urls) {
  const encabezados = ['FOLIO','CURP','NUM_CONSTANCIA','TIPO','NOMBRE_CURSO','FECHA','INSTITUCION','URL_CONSTANCIA'];
  const primeraFila = hoja.getLastRow() > 0 ? hoja.getRange(1,1).getValue() : '';
  if (hoja.getLastRow() === 0 || primeraFila !== 'FOLIO') {
    if (hoja.getLastRow() > 0) hoja.insertRowBefore(1);
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  }
  actualizacion.forEach(function(c, i) {
    hoja.appendRow([folio, curp, i+1, c.tipo, c.nombre, c.fecha, c.institucion, urls[i] || '']);
  });
}

// =========================================================================
// CORREO DE CONFIRMACIÓN
// Sin puntajes — con relación de documentos con hipervínculos
// =========================================================================
function enviarCorreoConfirmacion(formData, antiguedad, curp, folio, urls) {
  try {
    const correo  = formData.datosIdentificacion.correo;
    const nombre  = formData.datosPersonales.nombreNatural;
    const funcion = formData.datosLaborales.funcion;
    const nivel   = formData.formacion.nivel || '—';
    const promedio = formData.formacion.promedio ? formData.formacion.promedio : 'No registrado';
    const fecha   = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "dd 'de' MMMM 'de' yyyy, HH:mm"
    );

    // ── Antigüedad ────────────────────────────────────────────────────────
    const antSegey  = antiguedad.segey.anios + ' año(s) ' + antiguedad.segey.meses + ' mes(es)';
    const antFuncion = antiguedad.funcionEspec.anios + ' año(s) ' + antiguedad.funcionEspec.meses + ' mes(es)';

    // ── Relación de documentos con hipervínculo ───────────────────────────
    function filaDoc(nombre, url) {
      if (!url) return '';
      return '<tr><td>📄 <a href="' + url + '" style="color:#970E48;text-decoration:none;font-weight:600;">' + nombre + '</a></td></tr>';
    }

    let filasDocIdentif = filaDoc('Identificación Oficial', urls.identificacion);
    let filasDocFormacion = [
      filaDoc('Título Profesional',            urls.titulo),
      filaDoc('Cédula Profesional',            urls.cedula),
      filaDoc('Certificado de Estudios',       urls.certificado),
      filaDoc('Boleta de Calificaciones',      urls.boleta),
      filaDoc('Curso Práctico Afín',           urls.cursoNinias),
      filaDoc('Certificación Oficial Afín',    urls.certifNinias)
    ].join('');

    let filasDocAntiguedad = urls.urlsPeriodos.map(function(url, i) {
      return filaDoc(urls.nombresPeriodos ? urls.nombresPeriodos[i] : ('Soporte Período ' + (i+1)), url);
    }).join('');

    let filasDocActualizacion = urls.urlsActualizacion.map(function(url, i) {
      return filaDoc(urls.nombresActualizacion ? urls.nombresActualizacion[i] : ('Constancia ' + (i+1)), url);
    }).join('');

    let filasDocCarta = filaDoc('Carta de Reconocimiento al Desempeño', urls.carta);

    // Agrupar secciones de documentos
    function seccionDocs(titulo, filas) {
      if (!filas) return '';
      return '<tr><td style="padding:6px 12px;background:#f5eef4;font-size:11px;font-weight:700;color:#970E48;letter-spacing:.05em;text-transform:uppercase;">' + titulo + '</td></tr>' + filas;
    }

    const todasFilasDocs =
      seccionDocs('Identificación', filasDocIdentif) +
      seccionDocs('Formación Académica', filasDocFormacion) +
      seccionDocs('Antigüedad y Soportes', filasDocAntiguedad) +
      (filasDocActualizacion ? seccionDocs('Actualización y Desarrollo', filasDocActualizacion) : '') +
      seccionDocs('Carta de Reconocimiento', filasDocCarta);

    const cuerpoHtml = `
<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<style>
  body{font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0}
  .c{max-width:660px;margin:30px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.12)}
  .h{background:#970E48;padding:22px 28px;border-bottom:4px solid #C2995C}
  .h h1{color:#fff;font-size:17px;margin:0;font-weight:bold;letter-spacing:.8px}
  .h p{color:#f0c8d8;font-size:12px;margin:5px 0 0}
  .b{padding:26px 28px;color:#4A4A4A;font-size:14px;line-height:1.65}
  .ok{background:#e8f5e9;border-left:4px solid #2e7d32;padding:13px 16px;border-radius:4px;margin-bottom:18px;font-weight:bold;color:#1b5e20}
  .folio-box{background:#970E48;color:#fff;font-size:28px;font-weight:900;letter-spacing:.14em;text-align:center;padding:18px;border-radius:8px;margin:18px 0;border:3px solid #C2995C}
  .folio-label{font-size:11px;font-weight:400;letter-spacing:.08em;opacity:.8;display:block;margin-bottom:5px}
  table{width:100%;border-collapse:collapse;margin:14px 0;font-size:13px}
  th{background:#970E48;color:#fff;padding:9px 12px;text-align:left;font-size:12px;letter-spacing:.5px}
  td{padding:8px 12px;border-bottom:1px solid #eee}
  tr:nth-child(even) td{background:#fafafa}
  .av{background:#FFF8E1;border-left:4px solid #C2995C;padding:12px 16px;border-radius:4px;font-size:13px;color:#5d4037;margin-top:18px}
  .pie{background:#970E48;color:#f0c8d8;text-align:center;padding:14px;font-size:11px;border-top:3px solid #C2995C}
</style></head><body>
<div class="c">
  <div class="h">
    <h1>SECRETARÍA DE EDUCACIÓN — SEGEY</h1>
    <p>Subsecretaría de Educación Básica &nbsp;|&nbsp; Proceso de Admisión PAAE 2026-2027</p>
  </div>
  <div class="b">
    <div class="ok">✅ Su expediente de registro ha sido recibido exitosamente.</div>
    <p>Estimado/a <strong>${nombre}</strong>,</p>
    <p>Su participación en el <strong>Proceso de Admisión PAAE 2026-2027</strong> ha sido registrada. Conserve su folio:</p>

    <div class="folio-box">
      <span class="folio-label">FOLIO DE REGISTRO</span>
      ${folio}
    </div>

    <table>
      <tr><th colspan="2">DATOS DE IDENTIFICACIÓN Y FORMACIÓN</th></tr>
      <tr><td><strong>Folio</strong></td><td><strong>${folio}</strong></td></tr>
      <tr><td><strong>CURP</strong></td><td>${curp}</td></tr>
      <tr><td><strong>RFC</strong></td><td>${formData.datosIdentificacion.rfc}</td></tr>
      <tr><td><strong>Función a la que concursa</strong></td><td>${funcion}</td></tr>
      <tr><td><strong>Nivel académico registrado</strong></td><td>${nivel}</td></tr>
      <tr><td><strong>Promedio general de estudios</strong></td><td>${promedio}</td></tr>
      <tr><td><strong>Antigüedad total en SEGEY</strong></td><td>${antSegey}</td></tr>
      <tr><td><strong>Antigüedad en la función a la que aspira</strong></td><td>${antFuncion}</td></tr>
      <tr><td><strong>Fecha y hora de registro</strong></td><td>${fecha}</td></tr>
    </table>

    <table>
      <tr><th>DOCUMENTOS CARGADOS EN SU EXPEDIENTE</th></tr>
      ${todasFilasDocs}
    </table>

    <div class="av"><strong>⚠️ Importante:</strong> Conserve este correo y su folio <strong>${folio}</strong> como comprobante de registro. Los documentos listados arriba están disponibles en los hipervínculos correspondientes. El expediente está sujeto a revisión en su cita de validación.</div>
    <p style="margin-top:18px">Atentamente,<br><strong>Subsecretaría de Educación Básica</strong><br>Secretaría de Educación del Gobierno del Estado de Yucatán (SEGEY)</p>
  </div>
  <div class="pie">© 2026 Subsecretaría de Educación Básica — SEGEY &nbsp;|&nbsp; Folio: ${folio} &nbsp;|&nbsp; Mensaje automático, no responda.</div>
</div>
</body></html>`;

    MailApp.sendEmail({
      to:       correo,
      subject:  'Folio ' + folio + ' — Confirmación de Registro PAAE 2026-2027 | SEGEY',
      htmlBody: cuerpoHtml
    });

  } catch(e) {
    Logger.log('Error correo confirmación: ' + e.toString());
  }
}
