import * as XLSX from 'xlsx';

/**
 * Quita la extensión del nombre de archivo.
 * "Ventas_2024.xls" -> "Ventas_2024"
 */
export function quitarExtension(nombreArchivo) {
  return nombreArchivo.replace(/\.[^/.]+$/, '');
}

/**
 * Quita extensiones de Excel de nombres de hojas.
 * Se aplica de forma repetida por si el nombre viene como "archivo.xls.xlsx".
 */
export function quitarExtensionHoja(nombreHoja) {
  return nombreHoja.replace(/(\.xlsm?|\.xltm?|\.xltx)+$/i, '');
}

/**
 * Sanea un nombre para que sea válido como nombre de hoja de Excel:
 * máximo 31 caracteres, sin los símbolos \ / ? * [ ] :
 * y sin duplicados dentro del mismo libro.
 */
export function sanearNombreHoja(nombre, nombresUsados) {
  let limpio = quitarExtensionHoja(String(nombre ?? ''))
    .replace(/[\\/?*[\]:]/g, '')
    .trim();

  if (!limpio) limpio = 'Hoja';
  limpio = limpio.substring(0, 31);

  let nombreFinal = limpio;
  let contador = 1;

  while (nombresUsados.has(nombreFinal)) {
    const sufijo = `_${contador}`;
    nombreFinal = `${limpio.substring(0, 31 - sufijo.length)}${sufijo}`;
    contador += 1;
  }

  nombresUsados.add(nombreFinal);
  return nombreFinal;
}

/**
 * Lee un archivo (File) del navegador y devuelve su libro de trabajo.
 * Soporta .xls, .xlsx y .xlsm.
 */
export async function leerLibro(archivo) {
  const buffer = await archivo.arrayBuffer();

  return XLSX.read(buffer, {
    type: 'array',
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    cellDates: true,
  });
}

/** Crea un libro de trabajo vacío. */
export function crearLibroVacio() {
  return XLSX.utils.book_new();
}

/** Convierte un libro de trabajo en un Blob .xlsx listo para descargar. */
export function libroABlob(libro) {
  const salida = XLSX.write(libro, {
    bookType: 'xlsx',
    type: 'array',
    cellStyles: true,
  });

  return new Blob([salida], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/** Dispara la descarga de un Blob con el nombre indicado. */
export function descargarBlob(blob, nombreArchivo) {
  const url = URL.createObjectURL(blob);
  const enlace = document.createElement('a');

  enlace.href = url;
  enlace.download = nombreArchivo;

  document.body.appendChild(enlace);
  enlace.click();
  document.body.removeChild(enlace);

  // Se revoca después de que el navegador haya iniciado la descarga.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function clonarCelda(celda) {
  if (!celda || typeof celda !== 'object') return celda;
  return { ...celda };
}

function getCell(ws, row, col) {
  return ws[XLSX.utils.encode_cell({ r: row, c: col })];
}

function setCell(ws, row, col, celda) {
  const ref = XLSX.utils.encode_cell({ r: row, c: col });

  if (celda == null) {
    delete ws[ref];
    return;
  }

  ws[ref] = clonarCelda(celda);
}

/**
 * Devuelve la última fila/columna realmente usada.
 */
function obtenerLimitesHoja(ws) {
  if (!ws || !ws['!ref']) {
    return { minR: 0, minC: 0, maxR: -1, maxC: -1 };
  }

  const rango = XLSX.utils.decode_range(ws['!ref']);

  return {
    minR: rango.s.r,
    minC: rango.s.c,
    maxR: rango.e.r,
    maxC: rango.e.c,
  };
}

/**
 * Borra las filas cuya primera columna está vacía.
 *
 * Equivale a la primera parte del Office Script:
 *   if (valores[i][0] === "") -> delete row
 *
 * Se recorre de abajo hacia arriba para conservar correctamente
 * los índices.
 */
function eliminarFilasConPrimeraColumnaVacia(ws) {
  const limites = obtenerLimitesHoja(ws);

  if (limites.maxR < 0) return 0;

  let filasEliminadas = 0;

  for (let r = limites.maxR; r >= limites.minR; r -= 1) {
    const celdaA = getCell(ws, r, 0);
    const valor = celdaA?.v;

    if (valor === '' || valor == null) {
      for (let c = limites.maxC; c >= limites.minC; c -= 1) {
        delete ws[XLSX.utils.encode_cell({ r, c })];
      }

      // Mover todas las filas inferiores una posición hacia arriba.
      for (let rr = r + 1; rr <= limites.maxR; rr += 1) {
        for (let c = limites.minC; c <= limites.maxC; c += 1) {
          const origen = getCell(ws, rr, c);
          setCell(ws, rr - 1, c, origen);
          delete ws[XLSX.utils.encode_cell({ r: rr, c })];
        }
      }

      filasEliminadas += 1;
    }
  }

  const nuevoMaxR = limites.maxR - filasEliminadas;

  if (nuevoMaxR < limites.minR) {
    delete ws['!ref'];
  } else {
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: limites.minR, c: limites.minC },
      e: { r: nuevoMaxR, c: limites.maxC },
    });
  }

  return filasEliminadas;
}

/**
 * Inserta columnas conservando los objetos de las celdas.
 */
function insertarColumnas(ws, indice, cantidad) {
  if (cantidad <= 0) return;

  const limites = obtenerLimitesHoja(ws);
  if (limites.maxR < 0) return;

  for (let r = limites.minR; r <= limites.maxR; r += 1) {
    for (let c = limites.maxC; c >= indice; c -= 1) {
      const origen = getCell(ws, r, c);
      setCell(ws, r, c + cantidad, origen);
      delete ws[XLSX.utils.encode_cell({ r, c })];
    }
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: limites.minR, c: limites.minC },
    e: { r: limites.maxR, c: limites.maxC + cantidad },
  });
}

/**
 * Elimina columnas conservando los objetos de las celdas.
 */
function eliminarColumnas(ws, indice, cantidad) {
  if (cantidad <= 0) return;

  const limites = obtenerLimitesHoja(ws);
  if (limites.maxR < 0) return;

  const finEliminacion = indice + cantidad - 1;

  for (let r = limites.minR; r <= limites.maxR; r += 1) {
    for (let c = indice; c <= limites.maxC; c += 1) {
      if (c + cantidad <= limites.maxC) {
        const origen = getCell(ws, r, c + cantidad);
        setCell(ws, r, c, origen);
      } else {
        delete ws[XLSX.utils.encode_cell({ r, c })];
      }
    }
  }

  // Limpia cualquier celda que haya quedado en las columnas eliminadas.
  for (let r = limites.minR; r <= limites.maxR; r += 1) {
    for (let c = Math.max(indice, limites.maxC - cantidad + 1); c <= limites.maxC; c += 1) {
      delete ws[XLSX.utils.encode_cell({ r, c })];
    }
  }

  const nuevoMaxC = limites.maxC - cantidad;

  if (nuevoMaxC < limites.minC) {
    delete ws['!ref'];
  } else {
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: limites.minR, c: limites.minC },
      e: { r: limites.maxR, c: nuevoMaxC },
    });
  }
}

/**
 * Ajusta !ref y elimina celdas sobrantes que queden fuera del rango.
 */
function normalizarRef(ws) {
  const limites = obtenerLimitesHoja(ws);

  if (limites.maxR < 0 || limites.maxC < 0) {
    delete ws['!ref'];
    return;
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: limites.maxR, c: limites.maxC },
  });
}

function establecerValor(ws, referencia, valor) {
  const celda = ws[referencia] ?? {};
  celda.v = valor;
  celda.t = typeof valor === 'number' ? 'n' : 's';
  delete celda.f;
  ws[referencia] = celda;
}

function establecerFormula(ws, referencia, formula, valorCalculado = undefined) {
  const celda = ws[referencia] ?? {};

  celda.t = 'n';
  celda.f = formula;

  // SheetJS no calcula fórmulas. Si no existe un valor previo,
  // dejamos el resultado como undefined para que Excel lo recalcule.
  if (valorCalculado !== undefined) {
    celda.v = valorCalculado;
  } else {
    delete celda.v;
  }

  ws[referencia] = celda;
}

function numeroCelda(ws, row, col) {
  const celda = getCell(ws, row, col);
  if (!celda) return null;

  if (typeof celda.v === 'number' && Number.isFinite(celda.v)) return celda.v;

  if (typeof celda.v === 'string') {
    const texto = celda.v.trim().replace(/,$/, '').replace(/\s/g, '');
    if (!texto) return null;

    // Admite números como 370824, 370.824 y 370.824,50.
    if (/^-?\d+(?:\.\d+)?(?:,\d+)?$/.test(texto)) {
      const normalizado = texto.includes(',')
        ? texto.replace(/\./g, '').replace(',', '.')
        : texto;
      const numero = Number(normalizado);
      return Number.isFinite(numero) ? numero : null;
    }
  }

  return null;
}

function calcularValorAgp(ws, fila) {
  const n = numeroCelda(ws, fila - 1, 13);
  return n !== null && n > 17000 ? 17822 : 11100;
}

function calcularMargenContable(ws, fila) {
  const columnas = [11, 10, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23];
  const valores = columnas.map((col) => numeroCelda(ws, fila - 1, col) ?? 0);

  return (
    valores[0] - valores[1] + valores[2] - valores[3] - valores[4] -
    valores[5] - valores[6] - valores[7] - valores[8] + valores[9] +
    valores[10] + valores[11]
  );
}

function calcularTotalesNumericos(ws, filaInicio, filaFin, columnas = []) {
  const totales = {};

  columnas.forEach((col) => {
    let total = 0;

    for (let fila = filaInicio; fila <= filaFin; fila += 1) {
      let valor = numeroCelda(ws, fila - 1, col);

      // S = Agp (columna 19 / índice 18)
      if (col === 18) {
        valor = calcularValorAgp(ws, fila);
      }

      // X = Margen Contable (columna 24 / índice 23)
      if (col === 23) {
        valor = calcularMargenContable(ws, fila);
      }

      // Y replica X (columna 25 / índice 24)
      if (col === 24) {
        valor = calcularMargenContable(ws, fila);
      }

      if (valor !== null && Number.isFinite(valor)) total += valor;
    }

    totales[col] = total;
  });

  return totales;
}

function normalizarTextoBusqueda(valor) {
  if (valor == null) return '';
  return String(valor)
    .trim()
    .replace(/,$/, '')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function normalizarEncabezado(valor) {
  return String(valor ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function encontrarColumnaPorEncabezado(ws, nombres, maxFilas = 10) {
  const limites = obtenerLimitesHoja(ws);
  if (limites.maxR < 0) return null;

  const buscados = new Set(nombres.map(normalizarEncabezado));
  const ultimaFila = Math.min(limites.maxR, maxFilas - 1);

  for (let r = 0; r <= ultimaFila; r += 1) {
    for (let c = 0; c <= limites.maxC; c += 1) {
      const valor = getCell(ws, r, c)?.v;
      if (buscados.has(normalizarEncabezado(valor))) {
        return { col: c, headerRow: r };
      }
    }
  }

  return null;
}

function obtenerSucursalPrimeraLinea(ws) {
  const limites = obtenerLimitesHoja(ws);
  if (limites.maxR < 0) return 'Consulta Nota de Venta';

  // Primero intenta una celda de la primera línea que contenga "Sucursal".
  for (let c = 0; c <= limites.maxC; c += 1) {
    const valor = getCell(ws, 0, c)?.v;
    if (valor == null || String(valor).trim() === '') continue;

    const texto = String(valor).trim();
    const match = texto.match(/sucursal\s*[:\-]?\s*(.+)$/i);
    if (match?.[1]?.trim()) return match[1].trim();
  }

  // Si la primera fila viene como: Sucursal | Nombre.
  if (limites.maxC >= 1) {
    const primero = normalizarEncabezado(getCell(ws, 0, 0)?.v);
    if (primero === 'sucursal') {
      const segundo = getCell(ws, 0, 1)?.v;
      if (segundo != null && String(segundo).trim()) return String(segundo).trim();
    }
  }

  // Como último recurso, usa el primer valor no vacío de la primera línea.
  for (let c = 0; c <= limites.maxC; c += 1) {
    const valor = getCell(ws, 0, c)?.v;
    if (valor != null && String(valor).trim()) return String(valor).trim();
  }

  return 'Consulta Nota de Venta';
}

function obtenerNombreSucursalSeguro(nombre) {
  return String(nombre || 'Sucursal')
    .replace(/[\\/?*\[\]:]/g, '')
    .trim()
    .substring(0, 31) || 'Sucursal';
}

/**
 * Equivalente al Office Script entregado por el usuario.
 *
 * IMPORTANTE:
 * - Se ejecuta SOLO sobre las hojas principales.
 * - La hoja/libro adicional se agrega después y nunca pasa por esta función.
 * - La hoja "NV" se conserva sin transformación porque es la tabla de consulta
 *   utilizada por XLOOKUP.
 */
export function transformarHojaSegunScript(ws) {
  if (!ws || !ws['!ref']) {
    return;
  }

  // 1. Eliminar filas incompletas.
  eliminarFilasConPrimeraColumnaVacia(ws);

  if (!ws['!ref']) return;

  // 2. El "fin" posterior a la eliminación.
  const limitesDespuesFilas = obtenerLimitesHoja(ws);
  const fin = limitesDespuesFilas.maxR + 1; // Excel usa número de fila 1-based.

  // 3. Misma estructura de columnas del Office Script.
  //
  // A:A insert right
  insertarColumnas(ws, 0, 1);
  //
  // K:K insert right
  insertarColumnas(ws, 10, 1);
  //
  // M:N delete left
  eliminarColumnas(ws, 12, 2);
  //
  // O:O delete left
  eliminarColumnas(ws, 14, 1);
  //
  // Q:Q insert right
  insertarColumnas(ws, 16, 1);

  // 4. Encabezados.
  establecerValor(ws, 'A1', 'No');
  establecerValor(ws, 'K1', 'Costos');
  establecerValor(ws, 'Q1', 'Combustible');
  establecerValor(ws, 'T1', 'ADM');

  // Quita el relleno del encabezado, equivalente a clear() del fill.
  const encabezado = obtenerLimitesHoja(ws);
  for (let c = encabezado.minC; c <= encabezado.maxC; c += 1) {
    const celda = getCell(ws, 0, c);
    if (celda?.s?.fill) {
      celda.s = { ...celda.s };
      delete celda.s.fill;
    }
  }

  // 5. Datos/autollenado.
  if (fin >= 2) {
    establecerValor(ws, 'A2', 1);

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerValor(ws, `A${fila}`, fila - 1);
    }

    establecerValor(ws, 'Q2', 10000);
    for (let fila = 3; fila <= fin; fila += 1) {
      establecerValor(ws, `Q${fila}`, 10000);
    }

    establecerValor(ws, 'P2', 27000);
    for (let fila = 3; fila <= fin; fila += 1) {
      establecerValor(ws, `P${fila}`, 27000);
    }
  }

  // 6. Fórmulas por fila.
  establecerValor(ws, 'S1', 'Agp');

  if (fin >= 2) {
    establecerFormula(ws, 'S2', '=IF(N2>17000,17822,11100)');

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerFormula(
        ws,
        `S${fila}`,
        `=IF(N${fila}>17000,17822,11100)`,
      );
    }

    establecerValor(ws, 'X1', 'Margen Contable');
    establecerFormula(
      ws,
      'X2',
      '=L2-K2+N2-O2-P2-Q2-R2-S2-T2+U2+V2+W2',
    );

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerFormula(
        ws,
        `X${fila}`,
        `=L${fila}-K${fila}+N${fila}-O${fila}-P${fila}-Q${fila}-R${fila}-S${fila}-T${fila}+U${fila}+V${fila}+W${fila}`,
      );
    }

    // Y2 = +X2 y autofill.
    establecerFormula(ws, 'Y2', '=X2');

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerFormula(ws, `Y${fila}`, `=X${fila}`);
    }
  }

  // 7. Cruce con NV.
  establecerValor(ws, 'AA1', 'Chasis');

  // El cruce de Chasis se realiza después, cuando la hoja de consulta ya
  // está incorporada al libro. Aquí solo se prepara el encabezado.

  // 8. Fila de totales generales.
  // Se escribe con valores calculados (no solo con fórmulas) para que el
  // total sea visible incluso en visores que no recalculan fórmulas de XLSX.
  const filaTotal = fin + 1;
  establecerValor(ws, `A${filaTotal}`, 'TOTAL');

  const columnasTotales = [];
  for (let c = 9; c <= 24; c += 1) columnasTotales.push(c);
  // Y también es numérica y corresponde a la réplica de X.
  columnasTotales.push(25);

  const totales = calcularTotalesNumericos(ws, 2, fin, columnasTotales);

  // Los totales se dejan como fórmulas SUM para que sean DINÁMICOS.
  // Si posteriormente se agregan filas dentro del rango en Excel,
  // Excel ajustará automáticamente la referencia de la suma.
  // Se conserva además el valor calculado actual como caché para que
  // el total sea visible incluso antes de que Excel recalcule el libro.
  columnasTotales.forEach((c) => {
    const col = XLSX.utils.encode_col(c);
    establecerFormula(
      ws,
      `${col}${filaTotal}`,
      `=SUM(${col}2:${col}${fin})`,
      totales[c] ?? 0,
    );
  });

  // IMPORTANTE: establecerValor() no amplía !ref automáticamente.
  // Extendemos explícitamente el rango para que la fila TOTAL forme parte
  // real de la hoja y Excel/SheetJS no la descarte al normalizar el rango.
  const limitesConTotal = obtenerLimitesHoja(ws);
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: {
      r: Math.max(limitesConTotal.maxR, filaTotal - 1),
      c: Math.max(limitesConTotal.maxC, 25),
    },
  });

  // 9. Formato monetario J:Z.
  // Se limpia una coma final que pueda venir desde el archivo de origen
  // (por ejemplo, "370824,") antes de convertir el valor a número.
  const rango = obtenerLimitesHoja(ws);
  for (let r = 0; r <= rango.maxR; r += 1) {
    for (let c = 9; c <= 25; c += 1) {
      const celda = getCell(ws, r, c);
      if (!celda) continue;

      const copia = { ...celda };
      const valorOriginal = copia.v;

      if (typeof valorOriginal === 'string') {
        const texto = valorOriginal.trim();
        const sinComaFinal = texto.replace(/,$/, '');

        if (sinComaFinal !== texto && /^-?\d+(?:\.\d+)?$/.test(sinComaFinal)) {
          copia.v = Number(sinComaFinal);
          copia.t = 'n';
        }
      }

      // Pesos chilenos: sin decimales y con separador de miles.
      // En configuración regional española/chilena se visualiza 370.824.
      copia.z = '#,##0';
      ws[XLSX.utils.encode_cell({ r, c })] = copia;
    }
  }

  normalizarRef(ws);
}

/**
 * Normaliza un nombre para compararlo sin importar espacios, guiones,
 * guiones bajos o extensiones de Excel.
 */
function normalizarNombreConsulta(nombre = '') {
  return String(nombre)
    .replace(/\.[^.]+$/, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

/**
 * Busca la hoja que contiene la consulta de nota de venta.
 * Se prioriza una hoja cuyo nombre corresponda a "consulta nota venta".
 */
export function encontrarHojaConsultaNotaVenta(libro) {
  const candidatas = [
    'consultanotaventa',
    'consultanotadeventa',
    'notaventa',
  ];

  return libro.SheetNames.find((nombre) =>
    candidatas.includes(normalizarNombreConsulta(nombre)),
  ) || null;
}

/**
 * Agrega/actualiza el cruce de Chasis en la última columna de cada reporte.
 *
 * H de cada fila del reporte -> E de la hoja de consulta -> M de la consulta.
 * Se usa INDEX + MATCH en lugar de XLOOKUP para evitar #¿NOMBRE? en
 * versiones de Excel que no soportan XLOOKUP/BUSCARX. La hoja de consulta nunca se modifica.
 */
export function aplicarCruceChasis(libro, nombreHojaConsulta) {
  if (!nombreHojaConsulta || !libro.Sheets[nombreHojaConsulta]) return;

  const consulta = libro.Sheets[nombreHojaConsulta];
  const limitesConsulta = obtenerLimitesHoja(consulta);
  if (limitesConsulta.maxR < 0) return;

  // Se conserva la regla solicitada originalmente: H del reporte se busca
  // en E de la consulta y se devuelve M. Se escribe el resultado como valor,
  // evitando #¿NOMBRE? por diferencias de idioma/versión de Excel.
  const mapaChasis = new Map();
  for (let r = 1; r <= limitesConsulta.maxR; r += 1) {
    const clave = normalizarTextoBusqueda(getCell(consulta, r, 4)?.v);
    if (!clave || mapaChasis.has(clave)) continue;
    const chasis = getCell(consulta, r, 12)?.v;
    if (chasis != null && String(chasis).trim() !== '') {
      mapaChasis.set(clave, chasis);
    }
  }

  libro.SheetNames.forEach((nombreHoja) => {
    if (nombreHoja === nombreHojaConsulta) return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const limites = obtenerLimitesHoja(ws);
    if (limites.maxR < 1) return;

    establecerValor(ws, 'AA1', 'Chasis');

    let finDatos = limites.maxR + 1;
    const ultimaCeldaA = getCell(ws, limites.maxR, 0);
    if (String(ultimaCeldaA?.v ?? '').trim().toUpperCase() === 'TOTAL') {
      finDatos -= 1;
    }

    for (let fila = 2; fila <= finDatos; fila += 1) {
      const stock = normalizarTextoBusqueda(getCell(ws, fila - 1, 7)?.v);
      const chasis = mapaChasis.get(stock) ?? 'No Encontrado';
      establecerValor(ws, `AA${fila}`, chasis);
    }

    const rangoActualizado = obtenerLimitesHoja(ws);
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(rangoActualizado.maxR, finDatos - 1), c: Math.max(rangoActualizado.maxC, 26) },
    });
  });
}

/**
 * Cruza el STOCK del reporte con el STOCK de la consulta y agrega la
 * sucursal encontrada en una columna AB.
 *
 * Se intenta localizar el encabezado "Stock" automáticamente. Si el reporte
 * no lo trae como encabezado, se usa H como respaldo porque es la columna
 * definida por el flujo anterior.
 */
export function aplicarCruceSucursalPorStock(libro, nombreHojaConsulta) {
  if (!nombreHojaConsulta || !libro.Sheets[nombreHojaConsulta]) return;

  // La hoja opcional debe conservarse exactamente como "notas de venta".
  // La columna Sucursal de cada reporte se obtiene mediante BUSCARV/VLOOKUP:
  //
  // =BUSCARV(H2;'notas de venta'!$E:$AJ;32;0)
  //
  // En el XLSX las fórmulas se almacenan con los nombres internos de Excel,
  // por eso SheetJS debe escribir VLOOKUP, con comas y FALSE/0. Excel lo
  // mostrará como BUSCARV según el idioma de la instalación.
  const hojaConsultaReal = nombreHojaConsulta;
  const nombreHojaFormula = hojaConsultaReal.replace(/'/g, "''");

  libro.SheetNames.forEach((nombreHoja) => {
    if (nombreHoja === hojaConsultaReal) return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const limites = obtenerLimitesHoja(ws);
    if (limites.maxR < 1) return;

    // AB = Sucursal.
    establecerValor(ws, 'AB1', 'Sucursal');

    let finDatos = limites.maxR + 1;
    const ultimaCeldaA = getCell(ws, limites.maxR, 0);

    // La fila TOTAL no debe recibir BUSCARV.
    if (String(ultimaCeldaA?.v ?? '').trim().toUpperCase() === 'TOTAL') {
      finDatos -= 1;
    }

    // El usuario pidió específicamente H del reporte contra E:AJ de
    // "notas de venta", devolviendo la columna 32 del rango (AJ).
    for (let fila = 2; fila <= finDatos; fila += 1) {
      establecerFormula(
        ws,
        `AB${fila}`,
        `=VLOOKUP(H${fila},'${nombreHojaFormula}'!$E:$AJ,32,FALSE)`,
      );
    }

    const rangoActualizado = obtenerLimitesHoja(ws);
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: {
        r: Math.max(rangoActualizado.maxR, finDatos - 1),
        c: Math.max(rangoActualizado.maxC, 27),
      },
    });
  });
}

/**
 * Carga el valor ADM indicado por el usuario en la columna T (ADM) de cada
 * hoja principal del libro. La hoja "notas de venta" y cualquier hoja que
 * no tenga estructura de reporte quedan fuera.
 *
 * El valor se escribe en todas las filas de datos, no en la fila TOTAL.
 * La fila TOTAL ya contiene =SUM(T2:Tn), por lo que Excel recalculará el
 * total automáticamente al abrir el archivo.
 */

/**
 * Ajusta el Flete (columna P) según la Sucursal obtenida en AB.
 *
 * Si la sucursal contiene alguno de estos nombres:
 *   paicavi, trebol, ohiggins, prat
 * el flete queda en $85.000.
 * En cualquier otra sucursal se mantiene en $27.000.
 *
 * Se escribe como fórmula para que el valor sea dinámico y se actualice
 * automáticamente si cambia la Sucursal (AB).
 * La hoja "notas de venta" queda excluida.
 */
export function aplicarFletePorSucursal(libro) {
  if (!libro) return;

  const sucursalesFlete85000 = ['paicavi', 'trebol', 'ohiggins', 'prat'];

  libro.SheetNames.forEach((nombreHoja) => {
    if (normalizarNombreConsulta(nombreHoja) === 'notasdeventa') return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const encabezadoT = String(getCell(ws, 0, 19)?.v ?? '').trim().toUpperCase();
    if (encabezadoT !== 'ADM') return;

    const limites = obtenerLimitesHoja(ws);
    if (limites.maxR < 1) return;

    let finDatos = limites.maxR + 1;

    // La última fila TOTAL no recibe el cálculo del Flete.
    if (String(getCell(ws, limites.maxR, 0)?.v ?? '').trim().toUpperCase() === 'TOTAL') {
      finDatos -= 1;
    }

    if (finDatos < 2) return;

    const condiciones = sucursalesFlete85000
      .map((nombre) => `ISNUMBER(SEARCH("${nombre}",AB2))`)
      .join(',');

    for (let fila = 2; fila <= finDatos; fila += 1) {
      const formula = `IF(OR(${condiciones.replace(/AB2/g, `AB${fila}`)}),85000,27000)`;

      // P = Flete. Se deja como fórmula para que Excel determine el valor
      // según la Sucursal de AB en cada fila.
      establecerFormula(ws, `P${fila}`, formula, undefined);

      const celda = getCell(ws, fila - 1, 15);
      if (celda) {
        celda.z = '#,##0';
        celda.t = 'n';
      }
    }
  });

  return libro;
}

export function aplicarValorAdmEnHojas(libro, valorAdm) {
  if (!libro || !Number.isFinite(Number(valorAdm))) return;

  const valor = Number(valorAdm);

  libro.SheetNames.forEach((nombreHoja) => {
    if (normalizarNombreConsulta(nombreHoja) === 'notasdeventa') return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const limites = obtenerLimitesHoja(ws);
    if (limites.maxR < 1) return;

    // Solo hojas transformadas: ADM debe estar en T1.
    const encabezadoT = String(getCell(ws, 0, 19)?.v ?? '').trim().toUpperCase();
    if (encabezadoT !== 'ADM') return;

    let finDatos = limites.maxR + 1;
    const ultimaCeldaA = getCell(ws, limites.maxR, 0);
    if (String(ultimaCeldaA?.v ?? '').trim().toUpperCase() === 'TOTAL') {
      finDatos -= 1;
    }

    for (let fila = 2; fila <= finDatos; fila += 1) {
      establecerValor(ws, `T${fila}`, valor);
      const celda = getCell(ws, fila - 1, 19);
      if (celda) {
        celda.z = '#,##0';
      }
    }

    // Mantiene la fórmula dinámica del total de ADM.
    if (String(ultimaCeldaA?.v ?? '').trim().toUpperCase() === 'TOTAL') {
      establecerFormula(ws, `T${finDatos + 1}`, `=SUM(T2:T${finDatos})`);
    }

    normalizarRef(ws);
  });
}

/**
 * Detecta el nombre de la sucursal leyendo la primera línea de la consulta.
 */
export function detectarNombreSucursal(libro, nombreHojaConsulta) {
  const ws = libro.Sheets[nombreHojaConsulta];
  if (!ws) return 'Sucursal';
  return obtenerNombreSucursalSeguro(obtenerSucursalPrimeraLinea(ws));
}

/**
 * Renombra la hoja de consulta al nombre de la sucursal detectado.
 * Los datos no se modifican: solamente cambia el nombre de la hoja.
 */
export function renombrarHojaConsultaPorSucursal(libro, nombreHojaConsulta) {
  if (!nombreHojaConsulta || !libro.Sheets[nombreHojaConsulta]) return nombreHojaConsulta;

  const nombreSucursal = detectarNombreSucursal(libro, nombreHojaConsulta);
  let nombreFinal = nombreSucursal;
  let contador = 1;

  while (
    libro.SheetNames.some((nombre) => nombre !== nombreHojaConsulta && nombre.toLowerCase() === nombreFinal.toLowerCase())
  ) {
    const sufijo = `_${contador}`;
    nombreFinal = `${nombreSucursal.substring(0, 31 - sufijo.length)}${sufijo}`;
    contador += 1;
  }

  const indice = libro.SheetNames.indexOf(nombreHojaConsulta);
  libro.SheetNames[indice] = nombreFinal;
  libro.Sheets[nombreFinal] = libro.Sheets[nombreHojaConsulta];
  delete libro.Sheets[nombreHojaConsulta];

  return nombreFinal;
}

/**
 * Transforma todas las hojas principales del libro.
 *
 * La hoja NV se excluye porque es la fuente del cruce y el script original
 * no modifica esa hoja cuando se ejecuta sobre la hoja activa de datos.
 */
export function transformarLibro(libro) {
  libro.SheetNames.forEach((nombreHoja) => {
    if (nombreHoja.trim().toUpperCase() === 'NV') return;

    transformarHojaSegunScript(libro.Sheets[nombreHoja]);
  });

  // Las fórmulas agregadas deben recalcularse al abrir el XLSX en Excel.
  libro.Workbook = libro.Workbook || {};
  libro.Workbook.CalcPr = {
    ...(libro.Workbook.CalcPr || {}),
    calcMode: 'auto',
    fullCalcOnLoad: true,
    forceFullCalc: true,
  };

  return libro;
}

/**
 * Copia todas las hojas de `libroOrigen` dentro de `libroDestino`.
 *
 * Si el origen tiene una sola hoja, usa el nombre del archivo sin extensión.
 * Si tiene varias, conserva el nombre de cada hoja y elimina extensiones de
 * Excel que pudieran estar presentes.
 */
export function agregarHojasDeLibro(
  libroDestino,
  libroOrigen,
  prefijo,
  nombresUsados,
  opciones = {},
) {
  const { conservarNombresOriginales = false } = opciones;

  libroOrigen.SheetNames.forEach((nombreHoja) => {
    const nombreHojaSinExtension = quitarExtensionHoja(nombreHoja);

    const base = conservarNombresOriginales
      ? nombreHojaSinExtension
      : nombreHojaSinExtension.trim().toUpperCase() === 'NV'
        ? 'NV'
        : libroOrigen.SheetNames.length > 1
          ? `${prefijo}_${nombreHojaSinExtension}`
          : quitarExtension(prefijo);

    const nombreFinal = sanearNombreHoja(base, nombresUsados);

    XLSX.utils.book_append_sheet(
      libroDestino,
      libroOrigen.Sheets[nombreHoja],
      nombreFinal,
    );
  });
}

/**
 * Agrega el libro opcional exactamente como hojas adicionales, sin aplicar
 * ninguna transformación de datos. Sus nombres se conservan (sin extensión
 * de Excel) para que una hoja llamada "NV" siga pudiendo ser usada por
 * cruce.
 */
export function agregarHojasAdicionalesSinTransformar(
  libroDestino,
  libroOrigen,
  nombresUsados,
) {
  const hojasAgregadas = [];

  libroOrigen.SheetNames.forEach((nombreHojaOrigen, indice) => {
    const hojaOrigen = libroOrigen.Sheets[nombreHojaOrigen];
    if (!hojaOrigen) return;

    // La PRIMERA hoja del archivo opcional siempre se incorpora como
    // "notas de venta". No se transforma ni se altera su contenido.
    let nombreDestino;
    if (indice === 0) {
      nombreDestino = 'notas de venta';
    } else {
      nombreDestino = quitarExtensionHoja(nombreHojaOrigen);
    }

    nombreDestino = sanearNombreHoja(nombreDestino, nombresUsados);

    XLSX.utils.book_append_sheet(libroDestino, hojaOrigen, nombreDestino);
    hojasAgregadas.push(nombreDestino);
    nombresUsados.add(nombreDestino);
  });

  return hojasAgregadas;
}
