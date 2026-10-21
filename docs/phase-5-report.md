# Fase 5 - Informe de resultados reales (incrementos 5.1 y 5.2)

Fecha: 2026-10-21.
Entorno: Windows 11 Enterprise LTSC (10.0.26100, x64), .NET SDK 8.0.426.
Comandos ejecutados y resultados **observados** (no estimados).

> La Fase 5 empieza por lo que puede romper lo ya entregado: la carga de paquetes hostiles y el
> contrato de errores tipados. El resto de la fase (documentos grandes, decisiones pendientes y
> benchmarks) queda para incrementos posteriores.

## Sondeo previo (2026-10-21, antes de escribir código)

El sondeo se hizo con un programa aparte, un proceso por caso, para que un fallo del proceso no
tumbara la suite. Un `.xlsx` válido con la parte de hoja sustituida por XML hostil:

| Caso | Resultado del SDK 3.5.1 |
| --- | --- |
| Documento normal | Se abre |
| DTD sin entidad | `System.Xml.XmlException`: *DTD is prohibited* |
| *Billion laughs* (8 niveles de expansión) | `System.Xml.XmlException`: *DTD is prohibited* |
| Entidad externa (`file:///C:/Windows/win.ini`) | `System.Xml.XmlException`: *DTD is prohibited* |
| Carácter de control inválido | `System.Xml.XmlException`: *invalid character* |
| 1 000 niveles de anidamiento | Se abre |
| 2 000 niveles de anidamiento | **El proceso muere**: `0xC00000FD` (`STATUS_STACK_OVERFLOW`) |
| 10 000 niveles de anidamiento | **El proceso muere**: `0xC00000FD` |

Dos conclusiones: el SDK ya bloquea los DTD (pero con una excepción cruda y en un momento perezoso),
y el anidamiento profundo **termina el proceso** con un error que no se puede capturar. Ese segundo
punto es una denegación de servicio con un archivo de un solo kilobyte.

## Alcance

### Incremento 5.1 - Carga de OOXML hostil

- `OoxmlXmlGuard` (en `OfficeEngine.OpenXml`): mide las partes XML del paquete **antes** de que las
  parsee el SDK, con un lector de flujo que no usa la pila.
- Rechazo explícito de las declaraciones de tipo de documento.
- Límite de profundidad nuevo y público: `OfficeEngineLimits.MaximumXmlDepth` (256 por omisión).
- El guardián se invoca desde la apertura de Word y de Excel.

### Incremento 5.2 - Corpus de archivos corruptos y errores tipados

- `OoxmlPartReader`: la materialización del DOM de una parte traduce las familias de excepciones
  del SDK y del lector de XML a `InvalidOfficeFileException`, conservando la excepción original.
- La apertura y la inspección inicial traducen las mismas familias.
- `CorruptPackageTests`: corpus determinista de mutaciones sobre un `.docx` y un `.xlsx` válidos.

La decisión completa está en `docs/adr/0039-hostile-ooxml-loading.md`.

## Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **938 superadas, 0 con error, 0 omitidas** (938 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase5/inc51` | **936 superadas**; cobertura en `artifacts/testresults/phase5/inc51/**/coverage.cobertura.xml`. |

Salida textual del resumen (la duración varía):

```
Correctas! - Con error:     0, Superado:   938, Omitido:     0, Total:   938
```

Las **926 pruebas anteriores siguen pasando**; las 12 nuevas están en `HostilePackageTests` (7),
`CorruptPackageTests` (3) y dos pruebas de límites en `OfficeEngineLimitsTests`.

## Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `OpenXml/OoxmlXmlGuard.cs` | 50 / 51 | 98,0 % | 86,7 % |
| `OpenXml/OoxmlPartReader.cs` | 7 / 7 | 100 % | 50,0 % |
| **Total de la solución** | 5846 / 6268 | 93,3 % | 82,0 % |

La línea sin cubrir del guardián es el salto final de un flujo que no termina con el cierre esperado.
Las ramas sin cubrir del lector de partes son las familias de excepciones que no se han provocado.

## Pruebas nuevas (12 casos)

| Archivo | Foco |
| --- | --- |
| `Security/HostilePackageTests.cs` (7) | Anidamiento profundo (5 000 niveles) en un `.docx` y en un `.xlsx`; declaración de tipo de documento; entidad externa; carácter de control inválido; límite de profundidad configurable; y un documento válido que se sigue abriendo. |
| `Security/CorruptPackageTests.cs` (3) | Corpus de 25 mutaciones por formato (truncados, bytes alterados, entradas eliminadas y entradas sustituidas): cada una falla con un error tipado o se abre; y el tamaño del corpus. |
| `Core/CoreTests.cs` (+2) | El límite de profundidad por omisión y su validación. |

## Hallazgos durante las pruebas y acciones tomadas

1. **El sondeo encontró una denegación de servicio, no un error.** Un `.xlsx` de un kilobyte con
   2 000 niveles de anidamiento terminaba el proceso con `STATUS_STACK_OVERFLOW`. Ninguna
   `try`/`catch` sirve: el error se produce en una pila que ya no se puede recuperar. La única
   solución es **no llegar a ese punto**, y eso exige medir antes de parsear.
2. **El SDK ya bloquea los DTD, pero con una excepción cruda y tarde.** El fallo ocurre al
   materializar el DOM de la parte, es decir, después de abrir el archivo, y como `XmlException`. El
   guardián lo convierte en un rechazo explícito, con el nombre de la parte, en el momento de abrir.
3. **Una parte renombrada también es XML.** El guardián decide por los **bytes** (el primer byte no
   blanco es `<`), no por la extensión: un archivo con la parte XML renombrada a `.bin` se mide
   igual.
4. **El corpus encontró cuatro fugas de excepciones crudas**: `InvalidDataException` e
   `InvalidOperationException` del SDK, que se escapaban de la apertura y de la inspección
   inicial, `ArgumentOutOfRangeException`, que se escapaba al materializar una parte, y otra
   `InvalidDataException` al leer la tabla de cadenas compartidas. Las cuatro están traducidas
   ahora, y también la materialización de las partes de estilos, de numeración y de
   encabezados y pies.
5. **El corpus no es idéntico entre ejecuciones.** El SDK genera identificadores de relación
   distintos en cada archivo, así que un mismo byte alterado cae en un sitio distinto en cada
   ejecución. Eso es una ventaja: el corpus encuentra fugas que una ejecución afortunada no
   encontraría, pero también significa que **no** es un *fuzzer* determinista, sino un conjunto de
   mutaciones estable. La prueba se ejecutó ocho veces seguidas para confirmar la estabilidad.
6. **El coste de la medición es imperceptible.** La suite completa tarda 20 s con el guardián y 20 s
   sin él: el lector de flujo del guardián es más barato que el parseo con DOM que el SDK hace
   después.

### Lo que el corpus dejó fuera

- El corpus **no** es un *fuzzer*: son 25 mutaciones por formato sobre un archivo válido. Un
  *fuzzer* de verdad (AFL, libFuzzer) es un proyecto aparte, con su propia infraestructura.
- Las mutaciones cubren el contenedor ZIP y el contenido de las partes, no el **esquema** de OOXML:
  un documento con un elemento desconocido se abre y se conserva (eso ya lo cubre la matriz).

## Qué **no** se ha implementado (y por tanto no se afirma)

- Las otras actividades de la Fase 5: documentos grandes y el modelo de almacenamiento y `async`
  (decisión 2 de riesgos), y los benchmarks reproducibles.
- Un límite de **tiempo** de apertura: `OfficeEngineLimits` sigue sin exponer un tiempo máximo,
  porque cancelarlo exigiría un modelo asíncrono (ADR 0004).
- La revisión completa de la conservación de partes no soportadas: la matriz la cubre por casos
  concretos, no de forma exhaustiva.


## Incremento 5.3 - Documentos grandes, modelo de almacénamiento y `async`

### Alcance

Cerrar la decisión 2 de `docs/risks-and-dependencies.md` con datos, y comprobar que la librería
sostiene un documento grande. Incluye:

- Decisión documentada: el modelo en memoria se mantiene, la API sigue siendo síncrona y la regla
  que activa las variantes `...Async` es el cambio del almacén.
- Límite real de memoria documentado: el que manda es el del tamaño expandido, no el del archivo.
- Pruebas con un documento de 10 000 párrafos y un libro de 10 000 filas, y con un paquete que se
  expande por encima del límite.

La decisión completa está en `docs/adr/0040-in-memory-storage-and-synchronous-api.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **945 superadas, 0 con error, 0 omitidas** (945 total). |

Las **938 pruebas anteriores siguen pasando**; las 7 nuevas están en `LargeDocumentTests` (3),
`ExcelPerformanceTests` (1), `WordRoundTripRegressionTests` (2) y `ExcelRoundTripRegressionTests` (1).

### Medición del coste real (entorno de desarrollo)

Medido con `OfficeEngine.Benchmarks` en Windows 11 con .NET 8, con 5 000 párrafos o filas:

| Escenario | Tiempo | Memoria asignada | Paquete |
| --- | --- | --- | --- |
| `word.create` | 503 ms | 8,8 MB | 15 KB |
| `word.open` | 82 ms | 7,9 MB | - |
| `word.roundtrip` | 84 ms | 8,3 MB | - |
| `excel.create` | 1 787 ms | 29,8 MB | 81 KB |
| `excel.open` | 170 ms | 15,3 MB | - |
| `excel.roundtrip` | 198 ms | 16,2 MB | - |

El paquete comprimido es pequeño porque el XML se comprime muy bien; el coste real está en el DOM y
en el trabajo por celda.

### Hallazgos durante las pruebas y acciones tomadas

1. **El límite que manda no es el del archivo.** El ZIP descomprimido, el DOM del SDK y el buffer
   propio conviven: `MaxTotalUncompressedSizeBytes` es el límite de memoria, y el DOM lo multiplica.
   Los tres límites por omisión son una política, no una promesa de memoria, y así se documentan.
2. **Abrir un documento grande es líneal en su tamaño**, y no hay forma de cancelar una apertura en
   curso: no se puede, con la API síncrona actual.
3. **Escribir muchas celdas sigue siendo superlíneal** (ADR 0022): el acceso a cada celda recorre
   la hoja. Ahora está medido, no supuesto.

### Que **no** se ha implementado (y por tanto no se afirma)

- Un almacén respaldado por disco o por segmentos, y con ello una API `...Async` real.
- Cancelacion de una apertura o un guardado en curso.
- Un límite de tiempo de apertura: exigiria el modelo asíncrono.

## Incremento 5.4 - Benchmarks reproducibles y regresión de round-trip

### Alcance

Medir el rendimiento en lugar de suponerlo, y vigilar que el round-trip no se rompa. Incluye:

- `OfficeEngine.Benchmarks`: proyecto de consola que mide seis escenarios, escribe una referencia
  (`benchmarks/baseline.txt`) y compara contra ella.
- Integracion continua que publica el informe de medición como artefacto.
- `DocumentGenerator` y las regresiónes de round-trip sobre documentos generados.
- `ExcelPerformanceTests`, que vigila el coste cuadratico de escribir celdas.

La decisión completa está en `docs/adr/0041-benchmarks-and-round-trip-regression.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **945 superadas, 0 con error, 0 omitidas** (945 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase5/inc54` | **945 superadas**; cobertura en `artifacts/testresults/phase5/inc54/**/coverage.cobertura.xml`. |
| `dotnet run --project src/OfficeEngine.Benchmarks -- --baseline benchmarks/baseline.txt` | Comparación con la referencia: 6 escenas entre 0,96x y 1,01x. Código de salida 0. |

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| **Total de la solucion** | 5856 / 6278 | 93,3 % | 82,1 % |

### Hallazgos durante las pruebas y acciones tomadas

1. **La primera medición encontro un fallo de rendimiento real.** Escribir 10 000 celdas de Excel
   asígnaba **1 745 MB** para un paquete de 81 KB: la busqueda de una cadena en la tabla de cadenas
   compartidas recorria la tabla entera y reconstruia el texto de cada entrada en cada comparación.
   Con un indice, la misma operacion asígna **29,8 MB**: 58 veces menos, y el tiempo baja de 2,42 s
   a 1,79 s.
2. **La regresión queda protegida.** `ExcelPerformanceTests` falla si escribir 3 000 celdas vuelve a
   superar los 250 MB asígnados. Vive en una coleccion sin paralelismo para que la medida le
   pertenezca.
3. **No se anadio ninguna dependencia.** El proyecto de medición usa solo la biblioteca y la BCL:
   `BenchmarkDotNet` seria una dependencia nueva sin justificacion (regla 9 del `AGENT.md`).
4. **La comparación no es una puerta de CI.** El rendimiento de un runner compartido varia
   demásíado para tumbar una compilacion: CI publica el informe y la referencia se actualiza a
   proposito.

### Que **no** se ha implementado (y por tanto no se afirma)

- Un *fuzzer* de verdad (AFL, libFuzzer): el corpus de 5.2 es determinista, no aleatorio.
- Un perfil de asígnaciones por operacion ni un informe de GC detallado: la medida es el tiempo y la
  memoria asignada.
- Una puerta de rendimiento en CI: el informe se publica, no bloquea.

## Limitaciones conocidas

- El guardián recorre el paquete una vez más: en un paquete con partes XML muy grandes, el tiempo
  de apertura crece en proporción al tamaño de esas partes.
- El límite de profundidad es una política, no una medida: un documento con más de 256 niveles se
  rechaza aunque el SDK pudiera abrirlo. Es una decisión consciente, porque a partir de 1 500
  niveles el proceso muere.
- Las excepciones internas del lector de XML pueden incluir fragmentos del contenido del documento:
  no están en el mensaje público, pero quien conserve la excepción interna puede verlos.
