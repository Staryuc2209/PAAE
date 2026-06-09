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
// Usa LockService para garantizar que dos registros simultáneos
// nunca generen el mismo folio.
// La hoja 'Control_Folios' debe existir con A1='ULTIMO_FOLIO' y B1=0
// =========================================================================
function generarFolio() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // esperar hasta 10 segundos si hay otro proceso
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    let hoja   = ss.getSheetByName('Control_Folios');

    // Crear la hoja de control si no existe
    if (!hoja) {
      hoja = ss.insertSheet('Control_Folios');
      hoja.getRange('A1').setValue('ULTIMO_FOLIO');
      hoja.getRange('B1').setValue(0);
    }

    const ultimo = Number(hoja.getRange('B1').getValue()) || 0;
    const nuevo  = ultimo + 1;

    // Escribir el nuevo contador ANTES de liberar el lock
    hoja.getRange('B1').setValue(nuevo);
    SpreadsheetApp.flush(); // forzar escritura inmediata en disco

    return PREFIJO_FOLIO + String(nuevo).padStart(4, '0');
  } finally {
    lock.releaseLock(); // siempre liberar, incluso si hay error
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

    // ── Validar RFC: solo persona física (13 chars) ────────────────────────
    const rfc = formData.datosIdentificacion.rfc.trim().toUpperCase();
    if (rfc.length !== 13) {
      return { exito: false, mensaje: 'El RFC de persona física debe tener exactamente 13 caracteres.' };
    }

    // ── Generar folio único (con protección anti-concurrencia) ─────────────
    const folio = generarFolio();

    // ── Normalizar fechas ──────────────────────────────────────────────────
    formData.periodos = formData.periodos.map(function(p) {
      return {
        fechaInicio: normalizarFecha(p.fechaInicio),
        fechaFin:    normalizarFecha(p.fechaFin),
        funcion:     p.funcion
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
    enviarCorreoConfirmacion(formData, calculoAntiguedad, puntajes, curp, folio);

    return {
      exito:  true,
      folio:  folio,
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
  const idxCURP = datos[0].indexOf('CURP');
  if (idxCURP === -1) return false;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][idxCURP]).trim().toUpperCase() === curp) return true;
  }
  return false;
}

// =========================================================================
// CREACIÓN DE CARPETAS EN DRIVE
// El nombre incluye el folio al inicio para facilitar la búsqueda
// =========================================================================
function crearCarpetasAspirante(folio, curp, nombreCompleto) {
  const carpetaRaiz    = DriveApp.getFolderById(CARPETA_RAIZ_ID);
  const nombreCarpeta  = folio + '_' + curp + '_' + nombreCompleto.replace(/ /g,'_').replace(/\//g,'-');
  const carpetaUsuario = carpetaRaiz.createFolder(nombreCarpeta);
  carpetaUsuario.createFolder('01_Identificacion');
  carpetaUsuario.createFolder('02_Formacion_Academica');
  carpetaUsuario.createFolder('03_Antiguedad_Soportes');
  carpetaUsuario.createFolder('04_Actualizacion_Desarrollo');
  carpetaUsuario.createFolder('05_Carta_Reconocimiento');
  return carpetaUsuario;
}

// =========================================================================
// GUARDADO DE ARCHIVOS PDF EN DRIVE
// =========================================================================
function guardarArchivos(carpetaAspirante, formData, curp, folio) {
  const resultado = {
    identificacion:    '',
    carta:             '',
    urlsPeriodos:      [],
    urlsActualizacion: []
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
  resultado.identificacion = pdf(
    map['01_Identificacion'],
    formData.archivosRaw.identificacion,
    pre + '_IDENTIFICACION.pdf'
  );

  // 02 Formación Académica
  const fAc = map['02_Formacion_Academica'];
  if (formData.archivosRaw.titulo)       pdf(fAc, formData.archivosRaw.titulo,       pre + '_TITULO.pdf');
  if (formData.archivosRaw.cedula)       pdf(fAc, formData.archivosRaw.cedula,       pre + '_CEDULA.pdf');
  if (formData.archivosRaw.certificado)  pdf(fAc, formData.archivosRaw.certificado,  pre + '_CERTIFICADO.pdf');
  if (formData.archivosRaw.boleta)       pdf(fAc, formData.archivosRaw.boleta,       pre + '_BOLETA.pdf');
  if (formData.archivosRaw.cursoNinias)  pdf(fAc, formData.archivosRaw.cursoNinias,  pre + '_CURSO_AFIN.pdf');
  if (formData.archivosRaw.certifNinias) pdf(fAc, formData.archivosRaw.certifNinias, pre + '_CERTIF_AFIN.pdf');

  // 03 Períodos de antigüedad
  formData.periodos.forEach(function(p, i) {
    const url = (formData.archivosRaw.periodosFiles && formData.archivosRaw.periodosFiles[i])
      ? pdf(map['03_Antiguedad_Soportes'], formData.archivosRaw.periodosFiles[i],
            pre + '_PERIODO_' + (i + 1) + '.pdf')
      : '';
    resultado.urlsPeriodos.push(url);
  });

  // 04 Actualización
  formData.actualizacion.forEach(function(c, i) {
    const url = (formData.archivosRaw.actualizacionFiles && formData.archivosRaw.actualizacionFiles[i])
      ? pdf(map['04_Actualizacion_Desarrollo'], formData.archivosRaw.actualizacionFiles[i],
            pre + '_ACTUALIZACION_' + (i + 1) + '.pdf')
      : '';
    resultado.urlsActualizacion.push(url);
  });

  // 05 Carta de Reconocimiento
  resultado.carta = pdf(
    map['05_Carta_Reconocimiento'],
    formData.archivosRaw.carta,
    pre + '_CARTA_RECONOCIMIENTO.pdf'
  );

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

  const total = pAntiguedad + pReconocimiento + pFormacion + pActualizacion;
  return {
    antiguedad:     pAntiguedad,
    reconocimiento: pReconocimiento,
    formacion:      pFormacion,
    actualizacion:  pActualizacion,
    total:          total
  };
}

// =========================================================================
// REGISTRO EN HOJAS DE CÁLCULO
// =========================================================================
function registrarAspirante(hoja, formData, antiguedad, puntajes, urls, folio) {
  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      'FOLIO',
      'FECHA_REGISTRO','AP_PATERNO','AP_MATERNO','NOMBRES',
      'NOMBRE_INSTITUCIONAL','NOMBRE_NATURAL',
      'CURP','RFC','CORREO','TELEFONO',
      'FUNCION','NIVEL_ACADEMICO','PROMEDIO',
      'ANIOS_SEGEY','MESES_SEGEY','PTJE_ANTIGUEDAD',
      'ANIOS_RANGO','MESES_RANGO','PTJE_RANGO',
      'ANIOS_FUNCION','MESES_FUNCION','PTJE_FUNCION',
      'PTJE_RECONOCIMIENTO','PTJE_FORMACION','PTJE_ACTUALIZACION','PTJE_TOTAL',
      'URL_IDENTIFICACION','URL_CARTA'
    ]);
  }
  hoja.appendRow([
    folio,
    new Date(),
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
    Number(formData.formacion.promedio) || 0,
    antiguedad.segey.anios,
    antiguedad.segey.meses,
    antiguedad.segey.puntaje,
    antiguedad.rango.anios,
    antiguedad.rango.meses,
    antiguedad.rango.puntaje,
    antiguedad.funcionEspec.anios,
    antiguedad.funcionEspec.meses,
    antiguedad.funcionEspec.puntaje,
    puntajes.reconocimiento,
    puntajes.formacion,
    puntajes.actualizacion,
    puntajes.total,
    urls.identificacion,
    urls.carta
  ]);
}

function registrarPeriodos(hoja, curp, folio, periodos, urls) {
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(['FOLIO','CURP','NUM_PERIODO','FECHA_INICIO','FECHA_FIN','FUNCION','URL_SOPORTE']);
  }
  periodos.forEach(function(p, i) {
    hoja.appendRow([folio, curp, i + 1, p.fechaInicio, p.fechaFin, p.funcion, urls[i] || '']);
  });
}

function registrarActualizacion(hoja, curp, folio, actualizacion, urls) {
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(['FOLIO','CURP','NUM_CONSTANCIA','TIPO','NOMBRE_CURSO','FECHA','INSTITUCION','URL_CONSTANCIA']);
  }
  actualizacion.forEach(function(c, i) {
    hoja.appendRow([folio, curp, i + 1, c.tipo, c.nombre, c.fecha, c.institucion, urls[i] || '']);
  });
}

// =========================================================================
// CORREO DE CONFIRMACIÓN AL ASPIRANTE
// Ahora incluye el folio en asunto, encabezado y cuerpo
// =========================================================================
function enviarCorreoConfirmacion(formData, antiguedad, puntajes, curp, folio) {
  try {
    const correo  = formData.datosIdentificacion.correo;
    const nombre  = formData.datosPersonales.nombreNatural;
    const funcion = formData.datosLaborales.funcion;
    const fecha   = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "dd 'de' MMMM 'de' yyyy, HH:mm"
    );

    const filaAct = puntajes.actualizacion > 0
      ? '<tr><td>Actualización y Desarrollo Profesional</td><td><strong>' +
        puntajes.actualizacion.toFixed(2) + ' pts.</strong></td></tr>'
      : '<tr><td>Actualización y Desarrollo Profesional</td><td>Sin constancias registradas — 0 pts.</td></tr>';

    const cuerpoHtml = `
<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<style>
  body{font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0}
  .c{max-width:650px;margin:30px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.12)}
  .h{background:#970E48;padding:24px 30px;border-bottom:4px solid #C2995C}
  .h h1{color:#fff;font-size:18px;margin:0;font-weight:bold;letter-spacing:1px}
  .h p{color:#f0c8d8;font-size:13px;margin:6px 0 0}
  .b{padding:28px 30px;color:#4A4A4A;font-size:14px;line-height:1.6}
  .ok{background:#e8f5e9;border-left:4px solid #2e7d32;padding:14px 18px;border-radius:4px;margin-bottom:20px;font-weight:bold;color:#1b5e20}
  .folio-box{background:#970E48;color:#fff;font-size:26px;font-weight:900;letter-spacing:.12em;text-align:center;padding:18px;border-radius:8px;margin:20px 0;border:3px solid #C2995C}
  .folio-label{font-size:11px;font-weight:normal;letter-spacing:.08em;opacity:.85;display:block;margin-bottom:4px}
  table{width:100%;border-collapse:collapse;margin:18px 0;font-size:13px}
  th{background:#970E48;color:#fff;padding:9px 12px;text-align:left;font-size:12px;letter-spacing:.5px}
  td{padding:8px 12px;border-bottom:1px solid #e0e0e0}
  tr:nth-child(even) td{background:#F5EEF4}
  .pt{background:#3d0820;color:#E6C98A;font-size:20px;font-weight:bold;text-align:center;padding:16px;border-radius:6px;margin:20px 0;border:2px solid #C2995C}
  .av{background:#FFF8E1;border-left:4px solid #C2995C;padding:12px 16px;border-radius:4px;font-size:13px;color:#5d4037;margin-top:20px}
  .pie{background:#970E48;color:#f0c8d8;text-align:center;padding:16px;font-size:11px;border-top:3px solid #C2995C}
</style></head><body>
<div class="c">
  <div class="h">
    <h1>SECRETARÍA DE EDUCACIÓN — SEGEY</h1>
    <p>Subsecretaría de Educación Básica &nbsp;|&nbsp; Proceso de Admisión PAAE 2026-2027</p>
  </div>
  <div class="b">
    <div class="ok">✅ Su expediente de registro ha sido recibido exitosamente.</div>
    <p>Estimado/a <strong>${nombre}</strong>,</p>
    <p>Su participación en el <strong>Proceso de Admisión PAAE 2026-2027</strong> ha sido registrada. Conserve su folio de registro:</p>
    <div class="folio-box">
      <span class="folio-label">FOLIO DE REGISTRO</span>
      ${folio}
    </div>
    <table>
      <tr><th colspan="2">DATOS DE IDENTIFICACIÓN</th></tr>
      <tr><td><strong>Folio</strong></td><td><strong>${folio}</strong></td></tr>
      <tr><td><strong>CURP</strong></td><td>${curp}</td></tr>
      <tr><td><strong>RFC</strong></td><td>${formData.datosIdentificacion.rfc}</td></tr>
      <tr><td><strong>Función a la que concursa</strong></td><td>${funcion}</td></tr>
      <tr><td><strong>Nivel académico registrado</strong></td><td>${formData.formacion.nivel}</td></tr>
      <tr><td><strong>Fecha y hora de registro</strong></td><td>${fecha}</td></tr>
    </table>
    <table>
      <tr><th colspan="2">PUNTAJE PRELIMINAR REGISTRADO</th></tr>
      <tr><td>Antigüedad SEGEY</td><td>${antiguedad.segey.anios} año(s) ${antiguedad.segey.meses} mes(es) — <strong>${antiguedad.segey.puntaje.toFixed(2)} pts.</strong></td></tr>
      <tr><td>Carta de Reconocimiento al Desempeño</td><td><strong>${puntajes.reconocimiento} pts.</strong></td></tr>
      <tr><td>Formación Académica</td><td><strong>${puntajes.formacion} pts.</strong></td></tr>
      ${filaAct}
    </table>
    <div class="pt">PUNTAJE TOTAL PRELIMINAR: ${puntajes.total.toFixed(2)} / 100 pts.</div>
    <div class="av"><strong>⚠️ Importante:</strong> Este puntaje es <em>preliminar</em> y está sujeto a revisión en su cita de validación de expediente físico. Conserve este correo y su folio <strong>${folio}</strong> como comprobante.</div>
    <p style="margin-top:20px">Atentamente,<br><strong>Subsecretaría de Educación Básica</strong><br>Secretaría de Educación del Gobierno del Estado de Yucatán (SEGEY)</p>
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
