# ADR 0031 — Configuración de impresión y página

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.8

## Contexto

Una hoja se imprime con una configuración propia: los márgenes de la página, el tamaño y la
orientación del papel, la escala o el ajuste a un número de páginas, y qué se imprime (líneas de la
cuadrícula, títulos de fila y columna) y cómo se coloca en la página.

En SpreadsheetML son tres elementos de la hoja:

- `<printOptions>` con `horizontalCentered`, `verticalCentered`, `gridLines` y `headings`.
- `<pageMargins>` con `left`, `right`, `top`, `bottom`, `header` y `footer`, en pulgadas.
- `<pageSetup>` con `paperSize`, `orientation`, `scale`, `fitToWidth` y `fitToHeight`.

Además, el ajuste a páginas solo se aplica si la hoja declara `<sheetPr><pageSetUpPr fitToPage="1"/>`.
Y hay un cuarto elemento, `<headerFooter>`, con los textos de encabezado y pie, que esta versión no
modela.

## Decisión

### 1. API

- `Worksheet.PageMargins`: `ExcelPageMargins` (izquierda, derecha, arriba, abajo, encabezado y pie,
  en pulgadas) o `null`.
- `Worksheet.PageSetup`: `ExcelPageSetup` (orientación, tamaño de papel, escala y ajuste a páginas) o
  `null`.
- `Worksheet.PrintOptions`: `ExcelPrintOptions` (centrado horizontal y vertical, líneas de la
  cuadrícula y títulos) o `null`.

Asignar `null` a cualquiera de las tres quita el elemento; los demás atributos del elemento se
conservan porque solo se cambian los modelados.

### 2. Unidades y límites

- Los márgenes van en pulgadas, entre 0 y 100.
- El tamaño de papel es el identificador de SpreadsheetML (`ExcelPaperSize`): Carta (1), Tabloide (3),
  Ledger (4), Legal (5),Ejecutivo (7), A3 (8), A4 (9), A5 (11), B4 (12) y B5 (13).
- La escala es un porcentaje entre 10 y 400.

### 3. Escala y ajuste a páginas son excluyentes

Excel no puede usar las dos cosas a la vez: la escala fija el porcentaje y el ajuste deja que el
motor decida cuántas páginas hacen falta. Por eso el constructor de `ExcelPageSetup` rechaza un
`ExcelPageSetup` con escala **y** con ajuste, en lugar de escribir un archivo con las dos.

Cuando se usa el ajuste, la hoja marca `pageSetUpPr/@fitToPage="1"`, que es lo que hace que Excel lo
tenga en cuenta. Al quitar el ajuste, la marca se pone a `false` en lugar de desaparecer, para no
tocar el resto de `<sheetPr>`.

### 4. Lo que no se modela se conserva

- El encabezado y el pie (`<headerFooter>`) no tienen API: se conservan al guardar.
- Los atributos no modelados de `pageSetup` (por ejemplo `horizontalDpi` o `blackAndWhite`) se
  conservan al cambiar la configuración.

## Consecuencias

- La API pública nueva es `ExcelPageMargins`, `ExcelPageSetup`, `ExcelPageOrientation`,
  `ExcelPaperSize`, `ExcelPrintOptions` y las tres propiedades de `Worksheet`.
- La hoja producida puede contener `<printOptions>`, `<pageMargins>`, `<pageSetup>` y `<sheetPr>`.
- No se imprimen documentos: la configuración solo describe cómo se imprimiría la hoja.

## Verificación

- `ExcelPrintSettingsTests`: márgenes con round-trip, eliminación y límites; configuración de
  página con orientación, tamaño y escala; ajuste a páginas con la marca `fitToPage`; eliminación;
  valores inválidos; opciones de impresión con round-trip y eliminación; lectura de un archivo
  ajeno; libro liberado.
- `OoxmlValidation.CreateXlsxWithPrintSettings` construye el fixture con el SDK y
  `ReadXlsxPageMargins`, `ReadXlsxPageSetup` y `ReadXlsxPrintOptions` leen con el SDK.
