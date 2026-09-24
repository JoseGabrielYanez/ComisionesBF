import { useCallback, useRef, useState } from 'react';
import JSZip from 'jszip';
import {
  UploadCloud,
  FileSpreadsheet,
  X,
  FilePlus2,
  Loader2,
  CheckCircle2,
  AlertCircle,
  Lock,
  Download,
} from 'lucide-react';
import {
  leerLibro,
  crearLibroVacio,
  libroABlob,
  descargarBlob,
  agregarHojasDeLibro,
  agregarHojasAdicionalesSinTransformar,
  transformarLibro,
  aplicarCruceChasis,
  encontrarHojaConsultaNotaVenta,
  detectarNombreSucursal,
  renombrarHojaConsultaPorSucursal,
  aplicarCruceSucursalPorStock,
  aplicarValorAdmEnHojas,
  quitarExtension,
} from './utils/excelUtils';

const EXTENSIONES_VALIDAS = ['.xls', '.xlsx', '.xlsm'];

function esArchivoValido(archivo) {
  const nombre = archivo.name.toLowerCase();
  return EXTENSIONES_VALIDAS.some((ext) => nombre.endsWith(ext));
}

function esConsultaNotaVenta(archivo) {
  const nombre = quitarExtension(archivo?.name || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');

  return nombre === 'consultanotaventa' || nombre === 'consultanotadeventa';
}

function formatearTamano(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function App() {
  const [archivos, setArchivos] = useState([]);
  const [unirTodo, setUnirTodo] = useState(false);
  const [archivoExtra, setArchivoExtra] = useState(null);
  const [arrastrando, setArrastrando] = useState(false);
  const [procesando, setProcesando] = useState(false);
  const [error, setError] = useState('');
  const [exito, setExito] = useState('');
  const [valorAdm, setValorAdm] = useState('');

  const inputArchivosRef = useRef(null);
  const inputExtraRef = useRef(null);

  const agregarArchivos = useCallback((listaArchivos) => {
    setError('');
    setExito('');
    const candidatos = Array.from(listaArchivos);
    const validos = candidatos.filter(esArchivoValido);

    if (validos.length === 0) {
      setError('Selecciona archivos con extensión .xls o .xlsx.');
      return;
    }

    setArchivos((previos) => {
      const existentes = new Set(previos.map((a) => `${a.name}-${a.size}`));
      const sinDuplicados = validos.filter((a) => !existentes.has(`${a.name}-${a.size}`));
      return [...previos, ...sinDuplicados];
    });
  }, []);

  const onSeleccionarArchivos = (evento) => {
    agregarArchivos(evento.target.files);
    evento.target.value = '';
  };

  const onSoltarArchivos = (evento) => {
    evento.preventDefault();
    setArrastrando(false);
    agregarArchivos(evento.dataTransfer.files);
  };

  const quitarArchivo = (indice) => {
    setArchivos((previos) => previos.filter((_, i) => i !== indice));
  };

  const onSeleccionarExtra = (evento) => {
    const archivo = evento.target.files[0];
    evento.target.value = '';
    if (!archivo) return;

    if (!esArchivoValido(archivo)) {
      setError('El archivo adicional debe ser .xls o .xlsx.');
      return;
    }
    setError('');
    setArchivoExtra(archivo);
  };

  const manejarConversion = async () => {
    if (archivos.length === 0) {
      setError('Agrega al menos un archivo antes de convertir.');
      return;
    }

    setError('');
    setExito('');
    setProcesando(true);

    try {
      // Los archivos principales sí pasan por la transformación.
      // El archivo adicional NO se transforma: se agrega al final tal como
      // fue cargado, salvo la normalización de nombres de hoja (sin extensión).
      const librosPrincipales = await Promise.all(
        archivos.map(async (archivo) => {
          const libro = await leerLibro(archivo);
          transformarLibro(libro);

          return {
            nombre: quitarExtension(archivo.name),
            libro,
          };
        }),
      );

      const libroExtra = archivoExtra ? await leerLibro(archivoExtra) : null;

      if (unirTodo) {
        const libroFinal = crearLibroVacio();
        const nombresUsados = new Set();

        librosPrincipales.forEach(({ nombre, libro }) => {
          agregarHojasDeLibro(libroFinal, libro, nombre, nombresUsados);
        });

        // El valor ADM ingresado en el formulario se aplica a cada hoja
        // principal creada dentro de archivo_unido.xlsx. La hoja opcional
        // (notas de venta) queda completamente fuera de esta operación.
        if (valorAdm !== '') {
          aplicarValorAdmEnHojas(libroFinal, Number(valorAdm));
        }

        if (libroExtra) {
          // IMPORTANTE: el archivo adicional queda fuera de la
          // transformación y se incorpora al final.
          const hojasExtraAgregadas = agregarHojasAdicionalesSinTransformar(
            libroFinal,
            libroExtra,
            nombresUsados,
          );

          let hojaConsulta = encontrarHojaConsultaNotaVenta(libroFinal);
          if (esConsultaNotaVenta(archivoExtra) && hojasExtraAgregadas.length > 0) {
            // Si el archivo adicional se llama consulta_nota_venta, su primera
            // hoja se considera la fuente de consulta aunque internamente tenga
            // un nombre genérico como Hoja1.
            hojaConsulta = hojasExtraAgregadas[0];
          }

          if (hojaConsulta) {
            // La hoja del archivo opcional conserva el nombre "notas de venta".
            // Se utiliza como fuente, sin modificar sus datos.
            const nombreSucursal = detectarNombreSucursal(libroFinal, hojaConsulta);
            aplicarCruceChasis(libroFinal, hojaConsulta);
            aplicarCruceSucursalPorStock(libroFinal, hojaConsulta, nombreSucursal);
          }
        }

        descargarBlob(libroABlob(libroFinal), 'archivo_unido.xlsx');
        setExito('Listo: se descargó archivo_unido.xlsx');
      } else {
        const salidas = librosPrincipales.map(({ nombre, libro }) => {
          if (libroExtra) {
            const nombresUsados = new Set(libro.SheetNames);

            // IMPORTANTE: se agrega después de transformar el libro principal.
            const hojasExtraAgregadas = agregarHojasAdicionalesSinTransformar(
              libro,
              libroExtra,
              nombresUsados,
            );

            let hojaConsulta = encontrarHojaConsultaNotaVenta(libro);
            if (esConsultaNotaVenta(archivoExtra) && hojasExtraAgregadas.length > 0) {
              hojaConsulta = hojasExtraAgregadas[0];
            }

            if (hojaConsulta) {
              // La hoja del archivo opcional conserva el nombre "notas de venta".
              // Se utiliza como fuente, sin modificar sus datos.
              const nombreSucursal = detectarNombreSucursal(libro, hojaConsulta);
              aplicarCruceChasis(libro, hojaConsulta);
              aplicarCruceSucursalPorStock(libro, hojaConsulta, nombreSucursal);
            }
          }

          return { nombre: `${nombre}.xlsx`, blob: libroABlob(libro) };
        });

        if (salidas.length === 1) {
          descargarBlob(salidas[0].blob, salidas[0].nombre);
          setExito(`Listo: se descargó ${salidas[0].nombre}`);
        } else {
          const zip = new JSZip();
          salidas.forEach((salida) => zip.file(salida.nombre, salida.blob));
          const zipBlob = await zip.generateAsync({ type: 'blob' });
          descargarBlob(zipBlob, 'archivos_convertidos.zip');
          setExito(`Listo: se descargaron ${salidas.length} archivos en archivos_convertidos.zip`);
        }
      }
    } catch (err) {
      console.error(err);
      setError(`No se pudo completar la conversión: ${err.message}`);
    } finally {
      setProcesando(false);
    }
  };

  const cantidad = archivos.length;
  const resumen =
    cantidad === 0
      ? 'Agrega al menos un archivo para continuar.'
      : unirTodo
        ? `Se descargará archivo_unido.xlsx con ${cantidad} ${cantidad === 1 ? 'archivo' : 'archivos'}${archivoExtra ? ' y la hoja adicional al final' : ''}.`
        : cantidad === 1
          ? `Se descargará ${quitarExtension(archivos[0].name)}.xlsx.`
          : `Se descargará archivos_convertidos.zip con ${cantidad} archivos.`;

  const onCambiarAdm = (e) => setValorAdm(e.target.value.replace(/\D/g, '').slice(0, 12));
  const admVisible = valorAdm === '' ? '' : Number(valorAdm).toLocaleString('es-CL');

  const abrirSelector = () => inputArchivosRef.current?.click();
  const onTeclaZona = (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      abrirSelector();
    }
  };
  const onSalirZona = (e) => {
    if (!e.currentTarget.contains(e.relatedTarget)) setArrastrando(false);
  };

  const esquinas = [
    'left-0 top-0 border-l border-t',
    'right-0 top-0 border-r border-t',
    'bottom-0 left-0 border-b border-l',
    'bottom-0 right-0 border-b border-r',
  ];
  const foco =
    'focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-petrol';

  return (
    <div className="min-h-screen bg-white text-ink">
      <div className="mx-auto flex min-h-screen max-w-5xl flex-col px-6 py-5 sm:px-8">
        <header className="flex flex-wrap items-baseline justify-between gap-x-6 gap-y-1 border-b border-line pb-3 animate-rise">
          <h1 className="font-display text-3xl leading-none">Conversor de Excel</h1>
          <p className="flex items-center gap-2 text-xs text-mist">
            <Lock className="h-3 w-3 text-brass" strokeWidth={1.5} />
            Todo ocurre en tu navegador; los archivos no se envían a ningún servidor.
          </p>
        </header>

        <main className="flex flex-1 flex-col animate-rise" style={{ animationDelay: '80ms' }}>
          <div className="grid gap-x-12 gap-y-6 py-6 md:grid-cols-2">
            {/* Opciones (izquierda) */}
            <div className="space-y-5">
              <label className="flex cursor-pointer items-start justify-between gap-6">
                <span>
                  <span className="block text-sm font-semibold">Unir en un solo XLSX</span>
                  <span className="mt-0.5 block text-xs leading-5 text-mist">
                    Cada archivo pasa a ser una o varias hojas del mismo libro. Sin marcar, se
                    descarga uno por archivo.
                  </span>
                </span>
                <span className="relative mt-0.5 inline-flex flex-shrink-0">
                  <input
                    type="checkbox"
                    role="switch"
                    checked={unirTodo}
                    onChange={(e) => setUnirTodo(e.target.checked)}
                    className="peer sr-only"
                  />
                  <span className="h-5 w-9 rounded-full border border-slate-300 transition-colors peer-checked:border-petrol peer-checked:bg-petrol peer-focus-visible:ring-2 peer-focus-visible:ring-petrol/40" />
                  <span className="absolute left-0.5 top-0.5 h-4 w-4 rounded-full bg-slate-400 transition-all peer-checked:translate-x-4 peer-checked:bg-white" />
                </span>
              </label>

              <div className="border-t border-line pt-5">
                <label htmlFor="valor-adm" className="block text-sm font-semibold">
                  Valor ADM
                </label>
                <p className="mt-0.5 text-xs leading-5 text-mist">
                  Se carga en la columna ADM de cada hoja principal. La hoja notas de venta no se
                  modifica.
                </p>
                <div className="mt-2 flex items-baseline gap-2 border-b border-slate-300 pb-1 transition-colors focus-within:border-petrol">
                  <span className="font-display text-2xl text-brass">$</span>
                  <input
                    id="valor-adm"
                    type="text"
                    inputMode="numeric"
                    autoComplete="off"
                    value={admVisible}
                    onChange={onCambiarAdm}
                    placeholder="27.000"
                    className="tabular w-full bg-transparent font-display text-3xl outline-none placeholder:text-slate-300"
                  />
                </div>
                {!unirTodo && valorAdm !== '' && (
                  <p className="mt-1.5 text-xs text-mist">
                    Solo se aplica al unir los archivos en un solo XLSX.
                  </p>
                )}
              </div>

              <div className="border-t border-line pt-5">
                <p className="text-sm font-semibold">Hoja adicional (opcional)</p>
                <p className="mb-2.5 mt-0.5 text-xs leading-5 text-mist">
                  Se agrega al final del resultado, sin transformar.
                </p>
                {archivoExtra ? (
                  <div className="flex items-center justify-between gap-3 rounded-sm bg-wash px-3 py-2 text-sm">
                    <span className="flex min-w-0 items-center gap-2">
                      <FilePlus2 className="h-4 w-4 flex-shrink-0 text-brass" strokeWidth={1.5} />
                      <span className="truncate">{archivoExtra.name}</span>
                    </span>
                    <button
                      type="button"
                      onClick={() => setArchivoExtra(null)}
                      className={`flex-shrink-0 p-0.5 text-mist hover:text-ink ${foco}`}
                      aria-label="Quitar archivo adicional"
                    >
                      <X className="h-4 w-4" strokeWidth={1.5} />
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    onClick={() => inputExtraRef.current?.click()}
                    className={`inline-flex items-center gap-2 rounded-sm border border-slate-300 px-3 py-1.5 text-sm transition-colors hover:border-petrol hover:text-petrol ${foco}`}
                  >
                    <FilePlus2 className="h-4 w-4" strokeWidth={1.5} />
                    Elegir archivo
                  </button>
                )}
                <input
                  ref={inputExtraRef}
                  type="file"
                  accept=".xls,.xlsx,.xlsm"
                  onChange={onSeleccionarExtra}
                  className="hidden"
                />
              </div>
            </div>

            {/* Carga de archivos (derecha) */}
            <div
              role="button"
              tabIndex={0}
              aria-label="Elegir archivos para convertir"
              onClick={abrirSelector}
              onKeyDown={onTeclaZona}
              onDragOver={(e) => {
                e.preventDefault();
                setArrastrando(true);
              }}
              onDragLeave={onSalirZona}
              onDrop={onSoltarArchivos}
              className={`relative flex min-h-[190px] cursor-pointer flex-col items-center justify-center px-6 py-8 text-center outline-none transition-colors focus-visible:ring-2 focus-visible:ring-petrol ${
                arrastrando ? 'bg-petrol/5' : 'bg-wash hover:bg-slate-100'
              }`}
            >
              {esquinas.map((c) => (
                <span
                  key={c}
                  aria-hidden="true"
                  className={`pointer-events-none absolute h-3 w-3 transition-colors ${c} ${
                    arrastrando ? 'border-petrol' : 'border-brass'
                  }`}
                />
              ))}
              <UploadCloud className="mb-3 h-6 w-6 text-petrol" strokeWidth={1.25} />
              <p className="font-display text-2xl leading-tight">Arrastra tus archivos aquí</p>
              <p className="mt-1 text-xs text-mist">o haz clic para elegirlos · .xls, .xlsx</p>
              <input
                ref={inputArchivosRef}
                type="file"
                accept=".xls,.xlsx,.xlsm"
                multiple
                onChange={onSeleccionarArchivos}
                className="hidden"
              />
            </div>
          </div>

          {/* Lista de archivos */}
          <section className="border-t border-line py-4" aria-label="Archivos seleccionados">
            <h2 className="mb-2 text-sm font-semibold">
              Archivos seleccionados{' '}
              <span className="tabular font-normal text-mist">({cantidad})</span>
            </h2>
            {cantidad === 0 ? (
              <p className="text-xs text-mist">Aún no has agregado archivos.</p>
            ) : (
              <ul className="max-h-36 divide-y divide-line overflow-y-auto border-y border-line">
                {archivos.map((archivo, indice) => (
                  <li
                    key={`${archivo.name}-${archivo.size}-${indice}`}
                    className="flex items-center justify-between gap-4 py-1.5 text-sm"
                  >
                    <span className="flex min-w-0 items-center gap-2.5">
                      <FileSpreadsheet className="h-4 w-4 flex-shrink-0 text-brass" strokeWidth={1.5} />
                      <span className="truncate">{archivo.name}</span>
                      <span className="tabular flex-shrink-0 text-xs text-mist">
                        {formatearTamano(archivo.size)}
                      </span>
                    </span>
                    <button
                      type="button"
                      onClick={() => quitarArchivo(indice)}
                      className={`flex-shrink-0 p-0.5 text-mist hover:text-ink ${foco}`}
                      aria-label={`Quitar ${archivo.name}`}
                    >
                      <X className="h-4 w-4" strokeWidth={1.5} />
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </section>

          {/* Acción */}
          <div className="mt-auto border-t border-line pt-4">
            <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between md:gap-8">
              <p className="text-xs leading-5 text-mist">{resumen}</p>
              <button
                type="button"
                onClick={manejarConversion}
                disabled={procesando || cantidad === 0}
                className={`relative flex flex-shrink-0 items-center justify-center gap-2 overflow-hidden rounded-sm bg-petrol px-8 py-3 text-sm font-semibold text-white transition-colors hover:bg-petrol-dark disabled:cursor-not-allowed disabled:bg-slate-200 disabled:text-slate-500 ${foco}`}
              >
                {procesando ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" strokeWidth={1.75} />
                    Convirtiendo…
                    <span aria-hidden="true" className="absolute inset-x-0 bottom-0 h-0.5 overflow-hidden">
                      <span className="block h-full w-1/3 animate-sweep bg-brass" />
                    </span>
                  </>
                ) : (
                  <>
                    <Download className="h-4 w-4" strokeWidth={1.75} />
                    Convertir y descargar
                  </>
                )}
              </button>
            </div>

            {error && (
              <p role="alert" className="mt-3 flex items-start gap-2 text-sm text-danger">
                <AlertCircle className="mt-0.5 h-4 w-4 flex-shrink-0" strokeWidth={1.75} />
                {error}
              </p>
            )}
            {exito && (
              <p role="status" className="mt-3 flex items-start gap-2 text-sm text-ok">
                <CheckCircle2 className="mt-0.5 h-4 w-4 flex-shrink-0" strokeWidth={1.75} />
                {exito}
              </p>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
