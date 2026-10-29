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

// Devuelve el número de la última fila de datos (base 1), excluyendo la fila
// de totales situada dos filas después del último registro según el script.
function obtenerFinDatos(ws) {
  const limites = obtenerLimitesHoja(ws);
  if (limites.maxR < 0) return 0;
  const formulaTotal = getCell(ws, limites.maxR, 9)?.f;
  if (typeof formulaTotal === 'string' && /^SUM\(/i.test(formulaTotal)) {
    return Math.max(1, limites.maxR - 1);
  }
  return limites.maxR + 1;
}

/**
 * Borra las filas cuya primera columna está vacía.
 *
 * Equivale a la primera parte del Office Script:
 *   if (valores[i][0] === "") -> delete row
 *
 * Se compactan las filas conservadas una sola vez para evitar sobrescrituras.
 */
function eliminarFilasConPrimeraColumnaVacia(ws) {
  const limites = obtenerLimitesHoja(ws);

  if (limites.maxR < 0) return 0;
  // Calcula las filas que deben conservarse antes de mover cualquier celda.
  // El enfoque anterior desplazaba celdas varias veces cuando había filas
  // vacías consecutivas, sobrescribiendo registros del archivo fuente.
  const filasConservadas = [];
  for (let r = limites.minR; r <= limites.maxR; r += 1) {
    const valor = getCell(ws, r, 0)?.v;
    if (valor !== '' && valor != null) filasConservadas.push(r);
  }

  const filasEliminadas = limites.maxR - limites.minR + 1 - filasConservadas.length;
  const celdasFiltradas = new Map();
  filasConservadas.forEach((filaOrigen, indice) => {
    for (let c = limites.minC; c <= limites.maxC; c += 1) {
      const celda = getCell(ws, filaOrigen, c);
      if (celda) celdasFiltradas.set(XLSX.utils.encode_cell({ r: limites.minR + indice, c }), celda);
    }
  });

  for (let r = limites.minR; r <= limites.maxR; r += 1) {
    for (let c = limites.minC; c <= limites.maxC; c += 1) {
      delete ws[XLSX.utils.encode_cell({ r, c })];
    }
  }
  celdasFiltradas.forEach((celda, referencia) => { ws[referencia] = celda; });

  if (filasConservadas.length === 0) {
    delete ws['!ref'];
  } else {
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: limites.minR, c: limites.minC },
      e: { r: limites.minR + filasConservadas.length - 1, c: limites.maxC },
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

  // SheetJS almacena la fórmula sin el signo igual inicial.
  celda.f = String(formula).replace(/^=/, '');

  // SheetJS no calcula fórmulas. Guardamos el resultado conocido como caché
  // para que también se vea en visores que no recalculan el libro.
  if (valorCalculado !== undefined) {
    celda.v = valorCalculado;
    celda.t = typeof valorCalculado === 'number' ? 'n' : 'str';
  } else {
    delete celda.v;
    celda.t = 'n';
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

    let normalizado = null;
    if (/^-?\d{1,3}(?:\.\d{3})+$/.test(texto)) {
      normalizado = texto.replace(/\./g, '');
    } else if (/^-?\d{1,3}(?:\.\d{3})*,\d+$/.test(texto)) {
      normalizado = texto.replace(/\./g, '').replace(',', '.');
    } else if (/^-?\d+,\d+$/.test(texto)) {
      normalizado = texto.replace(',', '.');
    } else if (/^-?\d+(?:\.\d+)?$/.test(texto)) {
      normalizado = texto;
    }

    if (normalizado !== null) {
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
  const columnas = [11, 10, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22];
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

function crearMapaPorStock(ws, columnaResultado) {
  const limites = obtenerLimitesHoja(ws);
  const mapa = new Map();
  for (let fila = 1; fila <= limites.maxR; fila += 1) {
    const stock = normalizarTextoBusqueda(getCell(ws, fila, 4)?.v);
    const resultado = getCell(ws, fila, columnaResultado)?.v;
    if (!stock || resultado == null || String(resultado).trim() === '') continue;
    if (!mapa.has(stock)) mapa.set(stock, resultado);
  }
  return mapa;
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
 *   utilizada por BUSCARV.
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
  // Reserva U para Costo Kms y desplaza los bonos y márgenes al formato esperado.
  insertarColumnas(ws, 20, 1);

  // 4. Encabezados.
  establecerValor(ws, 'A1', 'No');
  establecerValor(ws, 'K1', 'Costos');
  establecerValor(ws, 'Q1', 'Combustible');
  establecerValor(ws, 'R1', 'Patente');
  establecerValor(ws, 'T1', 'ADM');
  establecerValor(ws, 'U1', 'Costo Kms');
  establecerValor(ws, 'V1', 'Bono Veh');
  establecerValor(ws, 'W1', 'Bono Santander');

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
    establecerFormula(ws, 'S2', '=IF(N2>17000,17822,11100)', calcularValorAgp(ws, 2));

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerFormula(
        ws,
        `S${fila}`,
        `=IF(N${fila}>17000,17822,11100)`,
        calcularValorAgp(ws, fila),
      );
    }

    establecerValor(ws, 'X1', 'Margen Contable');
    establecerValor(ws, 'Y1', 'Margen Vehículo');
    establecerFormula(
      ws,
      'X2',
      '=L2-K2+N2-O2-P2-Q2-R2-S2-T2+U2+V2+W2',
      calcularMargenContable(ws, 2),
    );

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerFormula(
        ws,
        `X${fila}`,
        `=L${fila}-K${fila}+N${fila}-O${fila}-P${fila}-Q${fila}-R${fila}-S${fila}-T${fila}+U${fila}+V${fila}+W${fila}`,
        calcularMargenContable(ws, fila),
      );
    }

    // Y2 = +X2 y autofill.
    establecerFormula(ws, 'Y2', '=X2', calcularMargenContable(ws, 2));

    for (let fila = 3; fila <= fin; fila += 1) {
      establecerFormula(ws, `Y${fila}`, `=X${fila}`, calcularMargenContable(ws, fila));
    }
  }

  // 7. Cruce con NV.
  insertarColumnas(ws, 27, 1);
  establecerValor(ws, 'AA1', 'Chasis');
  establecerValor(ws, 'AB1', 'Local');
  establecerValor(ws, 'Z1', 'Autos en Pago');
  establecerValor(ws, 'AC1', 'Acc');
  establecerValor(ws, 'AD1', 'Dif');

  // El cruce de Chasis se realiza después, cuando la hoja de consulta ya
  // está incorporada al libro. Aquí solo se prepara el encabezado.

  // 8. Totales: el script coloca SUM en J(fin+2) y rellena hasta Y.
  const filaTotal = fin + 2;
  const columnasTotales = Array.from({ length: 16 }, (_, indice) => indice + 9);
  const totales = calcularTotalesNumericos(ws, 2, fin, columnasTotales);
  for (let col = 9; col <= 24; col += 1) {
    const letra = XLSX.utils.encode_col(col);
    establecerFormula(
      ws,
      `${letra}${filaTotal}`,
      `=SUM(${letra}2:${letra}${fin})`,
      totales[col] ?? 0,
    );
  }

  // IMPORTANTE: establecerValor() no amplía !ref automáticamente.
  // Extendemos explícitamente el rango para que la fila TOTAL forme parte
  // real de la hoja y Excel/SheetJS no la descarte al normalizar el rango.
  const limitesConTotal = obtenerLimitesHoja(ws);
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: {
      r: Math.max(limitesConTotal.maxR, filaTotal - 1),
      c: Math.max(limitesConTotal.maxC, 30),
    },
  });

  // 9. Formato numérico con miles. XLSX guarda el separador como coma y Excel
  // lo muestra como punto en configuración regional chilena. También convierte
  // cadenas como "370.824," a números reales.
  const rango = obtenerLimitesHoja(ws);
  for (let r = 0; r <= rango.maxR; r += 1) {
    for (let c = 9; c <= 30; c += 1) {
      const celda = getCell(ws, r, c);
      if (!celda) continue;

      const copia = { ...celda };
      const valorOriginal = copia.v;

      if (typeof valorOriginal === 'string') {
        const texto = valorOriginal.trim().replace(/,$/, '').trim();
        let numeroNormalizado = null;

        if (/^-?\d{1,3}(?:\.\d{3})+$/.test(texto)) {
          numeroNormalizado = texto.replace(/\./g, '');
        } else if (/^-?\d+(?:\.\d+)?$/.test(texto)) {
          numeroNormalizado = texto;
        } else if (/^-?\d{1,3}(?:\.\d{3})*,\d+$/.test(texto)) {
          numeroNormalizado = texto.replace(/\./g, '').replace(',', '.');
        } else if (/^-?\d+,\d+$/.test(texto)) {
          numeroNormalizado = texto.replace(',', '.');
        }

        if (numeroNormalizado !== null && Number.isFinite(Number(numeroNormalizado))) {
          copia.v = Number(numeroNormalizado);
          copia.t = 'n';
        }
      }

      copia.z = '#,##0';
      ws[XLSX.utils.encode_cell({ r, c })] = copia;
    }
  }

  // Se conserva Z para Autos en Pago; Chasis, Local, Acc y Dif quedan a continuación.
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
 * Agrega/actualiza el cruce de Chasis en la columna AA.
 *
 * H de cada fila del reporte -> E de la hoja de consulta -> M de la consulta.
 * Busca H en la columna E de la hoja de consulta y devuelve la columna M,
 * equivalente a BUSCARV(H2;'notas de venta'!$E:$M;9;0). La hoja de consulta
 * nunca se modifica.
 */
export function aplicarCruceChasis(libro, nombreHojaConsulta) {
  if (!nombreHojaConsulta || !libro.Sheets[nombreHojaConsulta]) return;

  const consulta = libro.Sheets[nombreHojaConsulta];
  const limitesConsulta = obtenerLimitesHoja(consulta);
  if (limitesConsulta.maxR < 0) return;

  const nombreHojaFormula = nombreHojaConsulta.replace(/'/g, "''");
  const chasisPorStock = crearMapaPorStock(consulta, 12);

  libro.SheetNames.forEach((nombreHoja) => {
    if (nombreHoja === nombreHojaConsulta) return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const limites = obtenerLimitesHoja(ws);
    if (limites.maxR < 1) return;

    establecerValor(ws, 'AA1', 'Chasis');

    const finDatos = obtenerFinDatos(ws);

    for (let fila = 2; fila <= finDatos; fila += 1) {
      const stock = normalizarTextoBusqueda(getCell(ws, fila - 1, 7)?.v);
      const chasis = chasisPorStock.get(stock) ?? 'No Encontrado';
      establecerFormula(
        ws,
        `AA${fila}`,
        `=IFERROR(VLOOKUP(H${fila},'${nombreHojaFormula}'!$E:$M,9,FALSE),"No Encontrado")`,
        chasis,
      );
    }

    const rangoActualizado = obtenerLimitesHoja(ws);
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(rangoActualizado.maxR, finDatos - 1), c: Math.max(rangoActualizado.maxC, 26) },
    });
  });
}

/**
 * Completa Local usando el mismo cruce por stock que determina el nombre
 * de cada sucursal: H del reporte contra E:AJ de la consulta, devolviendo AJ.
 */
export function aplicarCruceLocal(libro, nombreHojaConsulta) {
  if (!nombreHojaConsulta || !libro.Sheets[nombreHojaConsulta]) return;

  const nombreHojaFormula = nombreHojaConsulta.replace(/'/g, "''");
  const localPorStock = crearMapaPorStock(libro.Sheets[nombreHojaConsulta], 35);
  libro.SheetNames.forEach((nombreHoja) => {
    if (nombreHoja === nombreHojaConsulta) return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    establecerValor(ws, 'AB1', 'Local');
    const finDatos = obtenerFinDatos(ws);
    for (let fila = 2; fila <= finDatos; fila += 1) {
      const stock = normalizarTextoBusqueda(getCell(ws, fila - 1, 7)?.v);
      const local = localPorStock.get(stock) ?? '';
      establecerFormula(
        ws,
        `AB${fila}`,
        `=IFERROR(VLOOKUP(H${fila},'${nombreHojaFormula}'!$E:$AJ,32,FALSE),"")`,
        local,
      );
    }

    const limites = obtenerLimitesHoja(ws);
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: limites.maxR, c: Math.max(limites.maxC, 27) },
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
 * Ajusta el Flete (columna P) según la sucursal encontrada en la consulta.
 *
 * Si la sucursal contiene alguno de estos nombres:
 *   paicavi, trebol, ohiggins, prat
 * el flete queda en $85.000.
 * En cualquier otra sucursal se mantiene en $27.000.
 *
 * Se escribe como fórmula para que el valor sea dinámico y se actualice
 * automáticamente si cambia la hoja de consulta.
 * La hoja "notas de venta" queda excluida.
 */
export function aplicarFletePorSucursal(libro, nombreHojaConsulta) {
  if (!libro) return;

  const sucursalesFlete85000 = ['paicavi', 'trebol', 'ohiggins', 'prat'];
  const nombreHojaFormula = String(nombreHojaConsulta || '').replace(/'/g, "''");
  if (!nombreHojaFormula || !libro.Sheets[nombreHojaConsulta]) return;
  const localPorStock = crearMapaPorStock(libro.Sheets[nombreHojaConsulta], 35);

  libro.SheetNames.forEach((nombreHoja) => {
    if (normalizarNombreConsulta(nombreHoja) === 'notasdeventa') return;

    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const encabezadoT = String(getCell(ws, 0, 19)?.v ?? '').trim().toUpperCase();
    if (encabezadoT !== 'ADM') return;

    const limites = obtenerLimitesHoja(ws);
    if (limites.maxR < 1) return;

    let finDatos = obtenerFinDatos(ws);

    if (finDatos < 2) return;

    for (let fila = 2; fila <= finDatos; fila += 1) {
      const stock = normalizarTextoBusqueda(getCell(ws, fila - 1, 7)?.v);
      const nombreLocal = normalizarNombreConsulta(localPorStock.get(stock) ?? '');
      const fleteCalculado = sucursalesFlete85000.some((nombre) => nombreLocal.includes(nombre))
        ? 85000
        : 27000;
      const busquedaSucursal = `IFERROR(VLOOKUP(H${fila},'${nombreHojaFormula}'!$E:$AJ,32,FALSE),"")`;
      const condiciones = sucursalesFlete85000
        .map((nombre) => `ISNUMBER(SEARCH("${nombre}",${busquedaSucursal}))`)
        .join(',');
      const formula = `=IF(OR(${condiciones}),85000,27000)`;

      // P = Flete. La sucursal se consulta directamente para no ocupar AB,
      // que pertenece a las columnas Acc/Dif del formato de salida.
      establecerFormula(ws, `P${fila}`, formula, fleteCalculado);

      const celda = getCell(ws, fila - 1, 15);
      if (celda) {
        celda.z = '#,##0';
        celda.t = 'n';
      }
    }
  });

  return libro;
}

/**
 * Renombra cada hoja de reporte según la sucursal asociada a sus stocks.
 * Busca H del reporte en E de la consulta y obtiene el nombre desde AJ
 * (columna 36 de la hoja). La hoja de consulta permanece intacta.
 */
export function renombrarHojasPorSucursal(libro, nombreHojaConsulta) {
  const consulta = libro?.Sheets?.[nombreHojaConsulta];
  if (!consulta || !nombreHojaConsulta) return libro;

  const limitesConsulta = obtenerLimitesHoja(consulta);
  if (limitesConsulta.maxR < 1 || limitesConsulta.maxC < 35) return libro;

  const sucursalPorStock = new Map();
  for (let fila = 1; fila <= limitesConsulta.maxR; fila += 1) {
    const stock = normalizarTextoBusqueda(getCell(consulta, fila, 4)?.v);
    const sucursal = getCell(consulta, fila, 35)?.v;
    if (!stock || sucursal == null || !String(sucursal).trim()) continue;
    if (!sucursalPorStock.has(stock)) sucursalPorStock.set(stock, sucursal);
  }

  const hojasReporte = libro.SheetNames.filter((nombre) => nombre !== nombreHojaConsulta);
  hojasReporte.forEach((nombreHoja) => {
    const ws = libro.Sheets[nombreHoja];
    if (!ws || !ws['!ref']) return;

    const finDatos = obtenerFinDatos(ws);
    let sucursalEncontrada = null;
    for (let fila = 2; fila <= finDatos; fila += 1) {
      const stock = normalizarTextoBusqueda(getCell(ws, fila - 1, 7)?.v);
      const sucursal = sucursalPorStock.get(stock);
      if (sucursal != null && String(sucursal).trim()) {
        sucursalEncontrada = String(sucursal).trim();
        break;
      }
    }
    if (!sucursalEncontrada) return;

    const nombreBase = sucursalEncontrada
      .replace(/\bventas\b/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const nombreSeguro = obtenerNombreSucursalSeguro(nombreBase);
    let nombreFinal = nombreSeguro;
    let contador = 1;
    while (
      libro.SheetNames.some(
        (nombre) => nombre !== nombreHoja && nombre.toLowerCase() === nombreFinal.toLowerCase(),
      )
    ) {
      const sufijo = `_${contador}`;
      nombreFinal = `${nombreSeguro.substring(0, 31 - sufijo.length)}${sufijo}`;
      contador += 1;
    }

    const indice = libro.SheetNames.indexOf(nombreHoja);
    libro.SheetNames[indice] = nombreFinal;
    libro.Sheets[nombreFinal] = ws;
    delete libro.Sheets[nombreHoja];
    if (libro.Workbook?.Sheets?.[indice]) {
      libro.Workbook.Sheets[indice].name = nombreFinal;
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

    let finDatos = obtenerFinDatos(ws);

    for (let fila = 2; fila <= finDatos; fila += 1) {
      establecerValor(ws, `T${fila}`, valor);
      const celda = getCell(ws, fila - 1, 19);
      if (celda) {
        celda.z = '#,##0';
      }
    }

    // Mantiene la fórmula dinámica del total de ADM.
    if (getCell(ws, limites.maxR, 9)?.f && /^SUM\(/i.test(getCell(ws, limites.maxR, 9).f)) {
      establecerFormula(
        ws,
        `T${finDatos + 2}`,
        `=SUM(T2:T${finDatos})`,
        valor * (finDatos - 1),
      );
    }

    normalizarRef(ws);
  });
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
