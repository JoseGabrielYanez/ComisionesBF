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

  return (
    <div className="min-h-screen bg-paper text-ink">
      <div className="mx-auto flex min-h-screen max-w-2xl flex-col justify-center px-6 py-12">
        <header className="mb-6">
          <h1 className="text-2xl font-semibold tracking-tight"> Hola Esteban, puedes cargar los archivos de comisiones</h1>
          <p className="mt-1 text-sm text-slate-600">
            Convierte tus archivos .xls a .xlsx, únelos en uno solo o súmales una hoja
            adicional — todo desde el navegador, sin subir nada a ningún servidor.
          </p>
        </header>

        <main className="divide-y divide-slate-200 rounded-xl border border-slate-200 bg-white">
          {/* Zona de carga de archivos */}
          <section className="p-6">
            <h2 className="mb-3 font-medium">Archivos a convertir</h2>

            <div
              onClick={() => inputArchivosRef.current?.click()}
              onDragOver={(e) => {
                e.preventDefault();
                setArrastrando(true);
              }}
              onDragLeave={() => setArrastrando(false)}
              onDrop={onSoltarArchivos}
              className={`cursor-pointer rounded-lg border-2 border-dashed p-8 text-center transition-colors ${
                arrastrando
                  ? 'border-accent bg-accent/5'
                  : 'border-slate-300 hover:border-accent/50'
              }`}
            >
              <UploadCloud className="mx-auto mb-3 h-7 w-7 text-slate-400" />
              <p className="font-medium">Arrastra tus archivos .xls aquí</p>
              <p className="text-sm text-slate-500">o haz clic para elegirlos desde tu equipo</p>
              <input
                ref={inputArchivosRef}
                type="file"
                accept=".xls,.xlsx,.xlsm"
                multiple
                onChange={onSeleccionarArchivos}
                className="hidden"
              />
            </div>

            {archivos.length > 0 && (
              <ul className="mt-4 space-y-2">
                {archivos.map((archivo, indice) => (
                  <li
                    key={`${archivo.name}-${archivo.size}-${indice}`}
                    className="flex items-center justify-between rounded-md border border-slate-200 px-3 py-2 text-sm"
                  >
                    <span className="flex min-w-0 items-center gap-2">
                      <FileSpreadsheet className="h-4 w-4 flex-shrink-0 text-accent" />
                      <span className="truncate">{archivo.name}</span>
                      <span className="flex-shrink-0 text-slate-400">
                        {formatearTamano(archivo.size)}
                      </span>
                    </span>
                    <button
                      type="button"
                      onClick={() => quitarArchivo(indice)}
                      className="ml-2 flex-shrink-0 rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                      aria-label={`Quitar ${archivo.name}`}
                    >
                      <X className="h-4 w-4" />
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </section>

          {/* Opción de unir archivos */}
          <section className="p-6">
            <label className="flex cursor-pointer items-start gap-3">
              <input
                type="checkbox"
                checked={unirTodo}
                onChange={(e) => setUnirTodo(e.target.checked)}
                className="mt-0.5 h-4 w-4 accent-accent"
              />
              <span>
                <span className="font-medium">Unir todos los archivos en un solo XLSX</span>
                <span className="block text-sm text-slate-500">
                  Cada archivo se convierte en una o varias hojas del mismo libro, con el
                  nombre del archivo original como prefijo. Si se deja sin marcar, cada
                  archivo se descarga convertido por separado.
                </span>
              </span>
            </label>
          </section>

          {/* Hoja adicional */}
          <section className="p-6">
            <h2 className="mb-1 font-medium">Hoja adicional (opcional)</h2>
            <p className="mb-3 text-sm text-slate-500">
              Sus hojas se agregan al final del resultado y quedan fuera de la
              transformación de datos.
            </p>

            {archivoExtra ? (
              <div className="flex items-center justify-between rounded-md border border-slate-200 px-3 py-2 text-sm">
                <span className="flex min-w-0 items-center gap-2">
                  <FilePlus2 className="h-4 w-4 flex-shrink-0 text-accent" />
                  <span className="truncate">{archivoExtra.name}</span>
                </span>
                <button
                  type="button"
                  onClick={() => setArchivoExtra(null)}
                  className="ml-2 flex-shrink-0 rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
                  aria-label="Quitar archivo adicional"
                >
                  <X className="h-4 w-4" />
                </button>
              </div>
            ) : (
              <button
                type="button"
                onClick={() => inputExtraRef.current?.click()}
                className="inline-flex items-center gap-2 rounded-md border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:border-accent/50 hover:text-accent"
              >
                <FilePlus2 className="h-4 w-4" />
                Elegir archivo adicional
              </button>
            )}
            <input
              ref={inputExtraRef}
              type="file"
              accept=".xls,.xlsx,.xlsm"
              onChange={onSeleccionarExtra}
              className="hidden"
            />
          </section>

          {/* Acción */}
          <section className="p-6">
            <button
              type="button"
              onClick={manejarConversion}
              disabled={procesando || archivos.length === 0}
              className="flex w-full items-center justify-center gap-2 rounded-md bg-accent px-4 py-2.5 font-medium text-white transition-colors hover:bg-accent-dark disabled:cursor-not-allowed disabled:bg-slate-300"
            >
              {procesando ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" />
                  Convirtiendo...
                </>
              ) : (
                'Convertir y descargar'
              )}
            </button>

            {error && (
              <p className="mt-3 flex items-start gap-2 text-sm text-red-600">
                <AlertCircle className="mt-0.5 h-4 w-4 flex-shrink-0" />
                {error}
              </p>
            )}
            {exito && (
              <p className="mt-3 flex items-start gap-2 text-sm text-emerald-600">
                <CheckCircle2 className="mt-0.5 h-4 w-4 flex-shrink-0" />
                {exito}
              </p>
            )}
          </section>
        </main>

        <footer className="mt-4 text-center text-xs text-slate-400">
          El procesamiento ocurre por completo en tu navegador; los archivos nunca se envían
          a un servidor.
        </footer>
      </div>
    </div>
  );
}
