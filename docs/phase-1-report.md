# Fase 1 — Informe de resultados reales

Fecha: 2026-10-16.
Entorno: Windows 10.0.26100 (x64), .NET SDK 8.0.426, runtime 8.0.32.
Comandos ejecutados y resultados **observados** (no estimados).

## Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Debug` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release` | **77 superadas, 0 con error, 0 omitidas** (77 total). Duración de la ejecución: ~150 ms. |

Salida textual del resumen de `dotnet test`:

```
Correctas! - Con error: 0, Superado: 77, Omitido: 0, Total: 77,
Duración: 150 ms - OfficeEngine.Tests.dll (net8.0)
```

El informe TRX se generó en `artifacts/testresults/phase1.trx`.

## Cobertura de la suite (77 pruebas)

| Área | Archivo | Qué verifica |
| --- | --- | --- |
| Núcleo | `Core/CoreTests.cs` | Límites válidos/ inválidos; diagnósticos (orden, severidad, snapshot, clear); validación de argumentos. |
| Seguridad | `ResourceLimitTests.cs` | Rechazo de no-ZIP, exceso de tamaño, número de entradas, tamaño total descomprimido y bombas ZIP; límites inválidos; jerarquía de excepciones tipadas. |
| Word | `Word/WordDocumentTests.cs` | Crear/abrir/guardar por archivo y stream; round-trip de párrafos; espacios significativos; guardados sucesivos; ciclo de vida de streams; `ObjectDisposedException`; validación de argumentos. |
| Word interop | `Word/WordInteropTests.cs` | Validación con `OpenXmlValidator`; contenido de `word/document.xml` crudo; lectura con el SDK; diagnóstico y **preservación de una tabla** tras round-trip. |
| Excel | `Excel/ExcelWorkbookTests.cs` | Crear/abrir/guardar; hoja por defecto; tipos texto/número/booleano/vacío; referencias inválidas; nombres de hoja (duplicados, longitud, caracteres, unicidad, renombrado); añadir/eliminar hojas; valor por defecto ausente no se crea al leer; `ObjectDisposedException`. |
| Excel interop | `Excel/ExcelInteropTests.cs` | Validación con `OpenXmlValidator`; `inlineStr` en `xl/worksheets/sheet1.xml`; lectura con el SDK; resolución de *shared strings*; **preservación de `mergeCells`**; lectura del valor en caché de una fórmula preservando la fórmula; detección de fila sin índice. |
| Documentación | `Documentation/ReadmeExamplesTests.cs` | Ejecuta el flujo de uso documentado en el README (Word y Excel) y el manejo de excepciones tipadas, para que los ejemplos no se desvíen de la API real. |

## Verificación independiente (no autorreferencial)

Las capacidades de la Fase 1 no se dan por buenas porque el contenedor ZIP se abra:

1. **Validación de esquema** con `DocumentFormat.OpenXml.Validation.OpenXmlValidator` sobre los
   paquetes producidos (`ProducedPackage_IsAValidWordprocessingDocument`,
   `ProducedPackage_IsAValidSpreadsheetDocument`).
2. **Lectura cruda** de entradas (`word/document.xml`, `xl/worksheets/sheet1.xml`,
   `xl/workbook.xml`) con `System.IO.Compression`.
3. **Lectura con el SDK** de documentos y libros producidos por OfficeEngine.
4. **Fixtures creados con el SDK** (tabla Word, *shared strings*, `mergeCells`, fórmula, fila
   malformada) leídos a través de OfficeEngine.

## Hallazgos durante las pruebas y acciones tomadas

1. Las tablas y las filas de prueba que construía el propio *fixture* no cumplían el esquema
   (faltaban `w:tblPr`/`w:tblGrid` y `w:r` en la fila). Se corrigieron **los fixtures**; no era
   un defecto de la librería. Esto confirmó que la validación con el SDK está detectando errores
   reales de esquema.
2. Se detectó que una fila sin el atributo `@r` (archivo malformado) hacía que la lectura
   devolviera `null` de forma silenciosa. Se cambió el comportamiento para lanzar
   `InvalidOfficeFileException` con un mensaje claro, y se añadió una prueba
   (`OpenedWorkbook_WithRowMissingIndex_...`). Preferimos un error explícito a una respuesta
   incorrecta silenciosa.

## Qué **no** se ha implementado (y por tanto no se afirma)

Ver `docs/compatibility-matrix.md`. En resumen:

- No hay formato de carácter/párrafo, estilos, listas, tablas editables, secciones,
  encabezados/pies, imágenes, hipervínculos ni búsqueda/reemplazo en Word.
- No hay estilos, formatos numéricos, fechas al escribir, rangos, nombres, validación de datos,
  filtros, gráficos ni escritura de fórmulas en Excel.
- No hay cálculo de fórmulas ni de campos.
- No hay exportación a PDF (LibreOffice no está instalado) ni soporte DOC/XLS.
- No hay soporte de documentos cifrados.

## Limitaciones conocidas de la Fase 1

- Documento completo en memoria; máximo por `OfficeEngineLimits.MaxFileSizeBytes` (256 MiB).
- I/O síncrona.
- Los diagnósticos se calculan al abrir y no se recalculan tras editar.
- Las cadenas se escriben como *inline strings*, no en tabla de cadenas compartidas.
