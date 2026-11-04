# Corrector TXT y reporte de mes

Una página con dos opciones: **CORRECTOR TXT** y **REPORTE DE MES**. Ambos procesos se ejecutan en el navegador; los archivos no se envían a servidores.

## Actualizar tu página

1. Descomprime el ZIP.
2. En tu repositorio, entra a **Add file → Upload files**.
3. Arrastra **todo el contenido extraído**, incluidas las carpetas `assets`, `lib` y `vendor`. Reemplaza los archivos anteriores. Sube el contenido, no el ZIP ni la carpeta que lo contiene.
4. Confirma con **Commit changes** en la rama utilizada por GitHub Pages.
5. Espera a que termine la publicación y abre la misma página. Si aparece la versión anterior, recarga con **Ctrl + F5**.

`index.html` debe quedar en la raíz del repositorio. No necesitas otro repositorio, instalar paquetes ni configurar claves. Si aún no tienes Pages activado: **Settings → Pages → Deploy from a branch → main → / (root) → Save**.

## Archivos incluidos

- `index.html`, `app.js`, `styles.css`: corrector TXT.
- `reporte.html`, `reporte.js`, `reporte.css`: reporte mensual.
- `lib/corrector.js` y `lib/report.js`: procesamiento.
- `assets/icon.svg`: icono.
- `vendor`: ExcelJS 4.4.0, JSZip 3.10.1 y sus licencias.
- `README.md` y `.nojekyll`.

Las bibliotecas están incluidas, sin CDN durante el uso. El ZIP no contiene bases de datos ni diarios reales.

## Reporte de mes

Selecciona **REPORTE DE MES**, carga la base `.xlsx`, revisa las exclusiones si hace falta y pulsa **Generar reporte**. Luego pulsa **Descargar reporte Excel**.

### Reglas

- Cuentas: `42120001`, `42120002`, `42410001`, `43120001` y `43120002`.
- Año 2026 según **Fecha comprobante**, tanto para incluidos como para excluidos.
- Exclusión parcial por nombre, ignorando tildes, mayúsculas, espacios y puntuación. No se aplican aproximaciones por errores ortográficos que puedan eliminar proveedores por accidente.
- Las aseguradoras se identifican por las palabras y nombres visibles en la página. **ONCOSALUD permanece incluido**. Añade cualquier aseguradora nueva que use otro nombre a esa lista.
- Las listas pueden editarse; los cambios no se guardan al salir o recargar.
- Se conservan las columnas anteriores a **Saldo EUR**, excepto **Importe EUR**. Se elimina **Saldo EUR** y todo lo que esté a su derecha.
- Los importes se muestran sin decimales, conservando la precisión real para los cálculos.

### Hojas

| Hoja | Contenido |
|---|---|
| TOTAL | Base limpia. |
| EXCLUIDOS | Registros excluidos por proveedor, solo cuentas seleccionadas de 2026, con el motivo. |
| RESUMEN | Una fila por proveedor con Saldo PEN y USD, más un total al final. |
| DATA | Base limpia más Mes, según Fecha comprobante, como fecha real con formato español `mmmm - yy`. |
| VCTO | Tabla dinámica nativa: proveedores en filas, meses en columnas y suma de Saldo PEN y USD. Con totales inferiores por mes y moneda; sin totales a la derecha. |

Los meses están en orden cronológico; aparecen solo los meses presentes en la base incluida. La fecha de vencimiento se conserva y no se usa para agrupar. Los registros de otras cuentas o años se omiten de todas las hojas; su cantidad aparece en la página.

RESUMEN agrupa por nombre exacto de proveedor, retirando espacios al principio y al final. No fusiona nombres distintos automáticamente.

La tabla dinámica incluye su caché y no requiere macros. Después de editar DATA en Excel, actualiza la tabla dinámica. Si agregas filas fuera del rango original, amplía su origen de datos. Para cambiar las reglas, genera otro reporte desde la base original.

### Entrada y límites

Una sola hoja debe tener los encabezados en la primera fila, incluyendo Proveedor, Cuenta principal, Fecha comprobante, Saldo PEN, Saldo USD y Saldo EUR. Se reconocen diferencias de mayúsculas, puntuación, tildes y espacios en los encabezados.

Hasta 20 MiB por archivo y 100 000 filas. Sin contraseña. Los importes deben ser números de Excel. Se rechazan fechas inválidas y fórmulas sin resultado guardado. Si no quedan registros incluidos, se muestra un aviso y no se descarga un reporte vacío.

## Corrector TXT

Mantiene la corrección anterior: cambia únicamente `10` por `01` en las columnas 86–87 del detalle, justo antes de `2026`, con el prefijo `00`: `00102026` pasa a `00012026`.

Espera una cabecera de 68 bytes que empieza por `P` y detalles de 107 bytes que empiezan por `6`. Conserva los demás bytes, espacios, codificación compatible con ASCII, BOM UTF-8 si existe y saltos CRLF, LF o CR. Verifica todo el archivo antes de descargar. No admite UTF-16 ni UTF-32. Hasta 20 archivos y 20 MiB por sesión; máximo 5 MiB por archivo.

## Privacidad

Publica solo los archivos de la herramienta. **No subas tus Excel o TXT de trabajo al repositorio.** `.gitignore` no controla el tratamiento de archivos dentro de una página.

La aplicación procesa los archivos en memoria y genera una descarga local. No utiliza historial, analítica, almacenamiento persistente del navegador ni servidor de procesamiento. La política de seguridad bloquea conexiones salientes del código. Los botones de limpiar liberan la selección y los resultados; las copias descargadas permanecen en tu equipo.

GitHub sirve la página y puede registrar datos habituales de la visita, como la dirección IP. El contenido de tus archivos no se envía.

## Verificación

46 pruebas automáticas del corrector, filtros y privacidad. Prueba de navegación, generación, descarga, limpieza y tamaño móvil en Edge. Con el ejemplo se compararon los registros conservados celda por celda y se conciliaron las sumas por proveedor y mes. Se verificaron las cinco hojas, XML, campos y registros de la caché de la tabla dinámica con un lector independiente. No se ha comprobado visualmente el resultado dentro de Excel de escritorio.


## Actualizar después de editar TOTAL

1. En el Excel generado, agrega, elimina o corrige registros en TOTAL. Conserva los encabezados y usa fechas de comprobante de 2026. Guarda el archivo.
2. En la página selecciona **Actualizar desde TOTAL**, carga ese Excel y pulsa **Actualizar desde TOTAL**.
3. Descarga el nuevo reporte. RESUMEN, DATA y VCTO se reconstruyen desde TOTAL, sin volver a filtrar cuentas o proveedores. Se conservan los valores y el orden de TOTAL; los formatos se generan de nuevo.

Puedes copiar una fila desde EXCLUIDOS a TOTAL. Si ambas filas tienen los mismos datos, la copia de EXCLUIDOS se retira automáticamente, una por cada coincidencia. Si además cambias sus datos, retírala tú de EXCLUIDOS para evitar que figure en ambas hojas. Las filas eliminadas de TOTAL no se archivan automáticamente en EXCLUIDOS. Si quieres conservarlas, muévelas a EXCLUIDOS antes de guardar.

Se genera un archivo nuevo con las cinco hojas. Otras hojas, anotaciones, formatos personalizados y fórmulas originales no se conservan: se utilizan sus valores guardados. Las fórmulas deben tener resultado calculado en Excel. No añadas una fila de totales manual en TOTAL.

VCTO conserva los totales inferiores y omite los totales a la derecha, tanto al generar como al actualizar.
