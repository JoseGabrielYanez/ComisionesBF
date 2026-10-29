# Conversor de Excel (XLS → XLSX)

Aplicación React + Vite que convierte archivos `.xls` a `.xlsx` directamente en el
navegador. Permite unir todos los archivos en un solo libro (con un checkbox) y
adjuntar un archivo adicional cuyas hojas se agregan al final del resultado.

## Requisitos

- Node.js 18 o superior
- npm

## Instalación

```bash
cd xls-a-xlsx-conversor
npm install
```

## Ejecutar en desarrollo

```bash
npm run dev
```

Abre la URL que indica la terminal (por defecto `http://localhost:5173`).

## Compilar para producción

```bash
npm run build
```

Los archivos listos para desplegar quedan en la carpeta `dist/` (se pueden servir
con cualquier hosting estático).

## Cómo funciona

1. **Archivos a convertir**: arrastra o selecciona uno o varios `.xls`/`.xlsx`.
2. **Unir todos los archivos en un solo XLSX** (checkbox): si se marca, todos los
   archivos se combinan en un único libro — cada archivo original se convierte en
   una o varias hojas, con su nombre de archivo como prefijo para evitar
   duplicados. Si no se marca, cada archivo se convierte y se descarga por
   separado (si son varios, se entregan juntos en un `.zip`).
3. **Hoja adicional (opcional)**: adjunta un archivo extra; sus hojas se agregan
   al final del libro resultante (o al final de cada archivo, si no se unieron).
4. Pulsa **Convertir y descargar**.

Todo el procesamiento se hace en el navegador con la librería
[SheetJS (`xlsx`)](https://www.npmjs.com/package/xlsx) y `jszip` para empaquetar
varias salidas; ningún archivo se sube a un servidor.

## Estructura del proyecto

```
src/
  App.jsx            Interfaz principal
  main.jsx           Punto de entrada
  index.css          Estilos (Tailwind)
  utils/
    excelUtils.js     Lectura, unión y descarga de libros de Excel
```


## Transformación de datos

Los archivos principales pasan automáticamente por una transformación equivalente
al Office Script indicado para este conversor:

- elimina filas con la primera columna vacía;
- inserta/elimina las columnas en la misma secuencia del script (la inserción opcional de Q permanece desactivada);
- crea `No`, `Costos`, `Combustible`, `ADM` y la columna `Costo Kms`;
- organiza bonos y márgenes en las columnas de la plantilla, y conserva `Autos en Pago`, `Acc` y `Dif`;
- rellena `No`, `Combustible` y el costo fijo;
- genera las fórmulas de `Agp`, `Margen Contable` y `Y`;
- agrega `Chasis` en `AA` y `Local` en `AB`, después de `Autos en Pago` en `Z`;
- completa `Local` con el stock de `H` buscado en `E:AJ` y devuelve la columna `AJ`;
- cuando se adjunta la consulta de nota de venta, busca `H` en `E:M` mediante `BUSCARV` y devuelve la columna `M`;
- renombra cada hoja principal con la sucursal obtenida al buscar su primer stock en `E` y devolver `AJ` (columna 36) desde la consulta; elimina la palabra `Ventas` del nombre;
- coloca las fórmulas de totales en `J(fin+2):Y(fin+2)`;
- limpia una coma final en valores monetarios cuando viene desde el archivo de origen;
- convierte valores numéricos de texto con coma final y aplica `#,##0`; Excel los muestra con punto de miles en configuración regional chilena.
- guarda también resultados calculados en las fórmulas para que se vean en Render y en visores que no recalculan el libro al abrirlo.

La hoja de consulta de nota de venta se mantiene íntegramente sin transformación.

### Archivo adicional

El archivo seleccionado en **Hoja adicional (opcional)** se procesa de forma
independiente y se agrega al final del resultado. **Nunca se aplica la
transformación de datos sobre ese archivo.** Sus nombres de hoja se conservan
(sin extensiones `.xls`, `.xlsx` o `.xlsm`) para permitir, por ejemplo, que una
hoja llamada `NV` siga siendo utilizada por las fórmulas de búsqueda.
