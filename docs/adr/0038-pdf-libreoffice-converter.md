# ADR 0038 — Conversión a PDF con LibreOffice headless

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 4, incrementos 4.1 y 4.2

## Contexto

El ADR 0005 decidió la estrategia: un contrato `IPdfConverter` en un proyecto aparte del núcleo y
un adaptador de motor externo detrás. La estrategia quedó a la espera de que hubiera un motor
verificado; el 2026-10-21 se comprobó que **LibreOffice 26.8.1.1** está instalado en
`C:\Program Files\LibreOffice\program\`, con licencia **MPL-2.0**, y que su modo headless convierte
`.docx` y `.xlsx` en PDF.

Quedaban por decidir las cosas que el enunciado del ADR 0005 señalaba: la forma del contrato, cómo se
exporta el estado en memoria, cómo se aísla el proceso, qué se hace cuando el motor falla y qué se
comprueba del PDF producido.

## Decisión

### 1. Dos proyectos, y el motor fuera del núcleo

- `OfficeEngine.Pdf` — solo contratos: `IPdfConverter`, `PdfSourceFormat`, `PdfConversionOptions`,
  `PdfConversionResult` y `PdfConversionException`. No depende de ningún motor y **no** lo referencia
  ni Word ni Excel.
- `OfficeEngine.Pdf.LibreOffice` — el adaptador: `LibreOfficePdfConverter` y
  `LibreOfficePdfConverterOptions`.

El núcleo (`OfficeEngine.Word`, `OfficeEngine.Excel`) no conoce el contrato: la conversión se pide
explícitamente, y el llamante decide qué bytes se convierten.

### 2. La conversión parte de los bytes que da el llamante

`IPdfConverter.Convert(Stream source, PdfSourceFormat format, PdfConversionOptions? options)`
recibe el documento como flujo. Para exportar el **estado actual** de un documento, el llamante lo
guarda en un `MemoryStream` y entrega ese `MemoryStream`: el adaptador nunca busca una copia antigua
en disco.

El formato se declara explícitamente (`WordDocument` o `ExcelWorkbook`) porque decide el filtro de
exportación y la extensión del archivo temporal.

### 3. El proceso se aísla por completo

Cada conversión usa:

- un directorio temporal **único**, bajo `%TEMP%\OfficeEngine.Pdf\<guid>`;
- un **perfil de usuario propio** de LibreOffice (`-env:UserInstallation`), que evita que la
  conversión se enganche a una instancia de LibreOffice que el usuario tenga abierta y que dos
  conversiones no choquen;
- un directorio de entrada con un archivo de nombre fijo (`source.docx` / `source.xlsx`), sin el
  nombre del documento del llamante;
- un directorio de salida propio.

El directorio se borra al terminar, salvo que se pida conservarlo (`KeepTemporaryFiles`), y el
borrado nunca tapa el resultado de la conversión.

### 4. Comprobaciones antes y después de convertir

- **Antes**: se localiza el ejecutable (la ruta indicada, las rutas habituales o el `PATH`) y se lee
  su versión. Si el motor no está, se lanza `PdfConversionException`; si la versión es anterior a la
  mínima (7.0 por omisión), también.
- **Durante**: el proceso se lanza con `UseShellExecute = false` y con un **tiempo máximo**
  (2 minutos por omisión). Si se pasa el tiempo, se mata el árbol de procesos y se lanza
  `PdfConversionException`.
- **Después**: se comprueban el **código de salida** (distinto de cero → error), la **existencia** del
  PDF, que no esté **vacío**, que no supere el **tamaño máximo** y que empiece por `%PDF-`. Un
  archivo que no sea un PDF es un error, no un resultado.

### 5. En Windows se usa el ejecutable de consola

`soffice.exe` es una aplicación gráfica que se desvincula: no espera y no devuelve el código de
salida. Por eso el adaptador busca `soffice.com`, su compañero de consola, en el mismo directorio.
Es lo que permite conocer el resultado de la conversión y su código de salida.

### 6. El motor es tolerante y el contrato lo refleja

LibreOffice abre casi cualquier cosa: unos bytes que no son un `.docx` se pueden imprimir como un
documento de texto. El adaptador **no** decide si el origen es válido; valida lo que devuelve. Un
documento corrupto puede, por tanto, producir un PDF válido: es el motor quien decide.

### 7. Límites y concurrencia

- El tamaño de la entrada y el de la salida tienen un máximo (64 MiB por omisión); superarlos lanza
  `OfficeResourceLimitExceededException`.
- Cada instancia del adaptador permite dos conversiones simultáneas (`MaximumConcurrentConversions`),
  lo que acota el número de procesos del motor sin serializar todo.
- El motor escribe ruido en la salida de error (por ejemplo, avisos del proveedor de scripts de
  Python). Es **información**, no un error: solo el código de salida y el archivo producido deciden
  si la conversión funcionó.

### 8. El contenido del documento no se registra

Los diagnósticos que devuelve la conversión son los mensajes del motor sobre la conversión. No
contienen el texto del documento, y el documento nunca se copia a un registro.

## Consecuencias

- La API pública nueva es `IPdfConverter`, `PdfSourceFormat`, `PdfConversionOptions`,
  `PdfConversionResult`, `PdfConversionException`, `LibreOfficePdfConverter` y
  `LibreOfficePdfConverterOptions`.
- El núcleo sigue sin depender de ningún motor, y el motor se puede sustituir por otro adaptador.
- La fidelidad del PDF depende de la versión de LibreOffice, de las fuentes instaladas y de la
  configuración del sistema: **no** hay equivalencia píxel a píxel con Microsoft Office.
- Las pruebas que necesitan el motor se marcan con `LibreOfficeFact` y **se omiten con un motivo
  explícito** cuando el motor no está instalado.

## Verificación

- `PdfContractTests`: validación de límites y tiempo de espera, y contenido de resultado y excepción.
- `LibreOfficePdfConverterTests` (integración): conversión de un `.docx` y de un `.xlsx` con el
  estado en memoria, comprobación de la cabecera `%PDF-`, conversión de bytes distintos que producen
  PDF distintos (el motor usa los bytes que recibe), documento corrupto que nunca devuelve algo que
  no sea un PDF, tiempo de espera agotado, ausencia de directorios temporales, flujo vacío, flujo
  nulo, formato desconocido, entrada por encima del límite y motor inexistente.
- `OoxmlValidation` y las pruebas de los proyectos `Word` y `Excel` verifican los paquetes que se
  convierten; el PDF se valida con su cabecera y con el validador del SDK cuando procede.
