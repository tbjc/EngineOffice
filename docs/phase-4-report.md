# Fase 4 - Informe de resultados reales (incrementos 4.1 y 4.2)

Fecha: 2026-10-21.
Entorno: Windows 11 Enterprise LTSC (10.0.26100, x64), .NET SDK 8.0.426, LibreOffice 26.8.1.1.
Comandos ejecutados y resultados **observados** (no estimados).

> La Fase 4 entrega la conversión a PDF: un contrato en el núcleo del producto y un adaptador de
> motor externo detrás. El motor se verificó antes de escribirlo, y las pruebas que lo necesitan
> se omiten con un motivo explícito cuando no está instalado.

## Alcance

### Incremento 4.1 - Contratos de conversión

- `OfficeEngine.Pdf`, un proyecto nuevo que no depende de ningún motor:
  - `IPdfConverter.Convert(Stream, PdfSourceFormat, PdfConversionOptions?)`.
  - `PdfSourceFormat` (`WordDocument`, `ExcelWorkbook`).
  - `PdfConversionOptions` (tamaño máximo de entrada y de salida, tiempo de espera, conservación
    de los temporales).
  - `PdfConversionResult` (bytes del PDF, diagnósticos del motor, duración).
  - `PdfConversionException` (código de salida y diagnósticos).

### Incremento 4.2 - Adaptador de LibreOffice headless

- `OfficeEngine.Pdf.LibreOffice`, con `LibreOfficePdfConverter` y sus opciones.
- Localización del ejecutable, comprobación de la versión, directorio temporal único por
  conversión, perfil de usuario aislado, tiempo de espera, límites de tamaño, comprobación del
  código de salida y del PDF producido, y borrado de los temporales.

La decisión completa está en `docs/adr/0038-pdf-libreoffice-converter.md`; la estrategia, en
`docs/adr/0005-pdf-conversion-strategy.md`.

## Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **926 superadas, 0 con error, 0 omitidas** (926 total). |
| `dotnet test ... --filter "FullyQualifiedName~Pdf"` | **17 superadas, 0 con error, 0 omitidas** (17 total). Duración: ~24 s. |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase4/inc41` | cobertura en `artifacts/testresults/phase4/inc41/**/coverage.cobertura.xml`. |

Salida textual del resumen (la duración varía):

```
Correctas! - Con error:     0, Superado:   926, Omitido:     0, Total:   926
```

Las **909 pruebas anteriores siguen pasando**; las 17 nuevas están en `PdfContractTests` (4) y
`LibreOfficePdfConverterTests` (13, de las que 7 necesitan el motor).

## Verificación del motor (2026-10-21)

| Comprobación | Resultado |
| --- | --- |
| Ejecutable | `C:\Program Files\LibreOffice\program\soffice.exe` |
| Versión | `LibreOffice 26.8.1.1 bf6ef2662401c7d6dce671ab4deb26ee0af3a7c1` |
| Licencia | MPL-2.0 (`license.txt`) |
| Headless | `.txt → .docx` (filtro *Office Open XML Text*) y `.docx → .pdf` (filtro `writer_pdf_Export`) |
| PDF producido | Cabecera `%PDF-` válida |

## Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Pdf/PdfConversionOptions.cs` | 14 / 14 | 100 % | 100 % |
| `Pdf/PdfConversionResult.cs` | 8 / 8 | 100 % | 100 % |
| `Pdf/PdfConversionException.cs` | 6 / 10 | 60,0 % | 100 % |
| `Pdf.LibreOffice/LibreOfficePdfConverter.cs` | 153 / 200 | 76,5 % | 64,1 % |
| `Pdf.LibreOffice/LibreOfficePdfConverterOptions.cs` | 4 / 4 | 100 % | 100 % |
| **Total de la solución** | 5757 / 6177 | 93,2 % | 81,9 % |

Las líneas sin cubrir del adaptador son las rutas que no se dan en Windows: la búsqueda del
ejecutable en el `PATH` y en las rutas de Unix, la lectura de una versión con un formato
inesperado y la construcción de la URI del perfil en un sistema sin estructura de unidad. Las dos
líneas sin cubrir de la excepción son el constructor con excepción interna y el `ToString` heredado.

## Pruebas nuevas (17 casos)

| Archivo | Foco |
| --- | --- |
| `Pdf/PdfContractTests.cs` (4) | Límites y tiempo de espera inválidos; valores por omisión; contenido del resultado y de la excepción. |
| `Pdf/LibreOfficePdfConverterTests.cs` (13) | Conversión de un `.docx` y de un `.xlsx` con el estado en memoria y cabecera `%PDF-`; dos contenidos distintos producen PDF distintos (el motor usa los bytes que recibe); un documento corrupto nunca devuelve algo que no sea un PDF; tiempo de espera agotado; no queda ningún directorio temporal; flujo vacío, nulo y formato desconocido; entrada por encima del límite; motor inexistente; ruta y versión del motor. |

Las siete pruebas que necesitan el motor se marcan con `LibreOfficeFact`, que las **omite con el
motivo explícito** *"LibreOffice is not installed in this environment."* cuando no está disponible.

## Hallazgos durante las pruebas y acciones tomadas

1. **En Windows hay que usar `soffice.com`.** `soffice.exe` es una aplicación gráfica que se
   desvincula: no espera al proceso ni devuelve el código de salida, y su salida estándar se
   pierde. El adaptador busca ahora su compañero de consola en el mismo directorio, que es lo que
   permite saber si la conversión funcionó.
2. **El motor es tolerante.** Un archivo que no es un `.docx` válido no falla: LibreOffice lo abre
   como texto y lo imprime. La prueba que esperaba un fallo se reescribió para comprobar lo que
   importa: que el adaptador **nunca** devuelve algo que no sea un PDF.
3. **El ruido de la salida de error no es un error.** El motor escribe avisos del proveedor de
   scripts de Python en la salida de error. La conversión se decide por el código de salida y por
   el archivo producido, nunca por el texto de la salida.
4. **El perfil de usuario es obligatorio.** Sin un perfil aislado por conversión, el motor se
   engancha a la instancia de LibreOffice que el usuario tenga abierta y no convierte nada.
5. **El documento no se copia a disco con su nombre.** El archivo temporal se llama `source.docx` o
   `source.xlsx`, en un directorio único por conversión, y se borra al terminar.

## Qué **no** se ha implementado (y por tanto no se afirma)

- **Equivalencia visual con Microsoft Office**: el PDF depende de la versión de LibreOffice, de las
  fuentes instaladas y de la configuración del sistema.
- **Selección de página, marcas de agua, cifrado del PDF, metadatos, Marcadores o formularios** del
  PDF: el adaptador solo produce el PDF que el motor genera.
- **Conversión de PDF a otro formato**, y conversión de `.doc`/`.xls` heredados: la Fase 6.
- **Un motor propio**: el cálculo y el renderizado siguen siendo de LibreOffice.

## Limitaciones conocidas

- La conversión necesita el motor instalado: sin él, `LibreOfficePdfConverter.Convert` lanza
  `PdfConversionException` y las pruebas de integración se omiten.
- Cada conversión arranca un proceso del motor: cuesta unos segundos y no es instantáneo.
- El motor decide el formato de página (A4, Letter) a partir de su configuración, no de la de
  OfficeEngine, salvo que el documento fije sus propios márgenes y tamaño de página.
