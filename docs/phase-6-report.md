# Fase 6 - Informe de resultados reales (incremento 6.1)

Fecha: 2026-10-21.
Entorno: Windows 11 Enterprise LTSC (10.0.26100, x64), .NET SDK 8.0.426, LibreOffice 26.8.1.1.
Comandos ejecutados y resultados **observados** (no estimados).

> La Fase 6 tiene dos mitades: la conversión de los formatos heredados con un motor externo, y el
> motor de fórmulas propio. Este informe recoge la primera; la segunda se mide en su propio
> incremento, con su matriz de funciones públicada.

## Alcance del incremento 6.1

Convertir los formatos heredados de Word y de Excel a sus equivalentes modernos. Incluye:

- `OfficeEngine.Legacy`, con `LegacyConverter`, `LegacyConverterOptions`, `LegacyConversionResult`,
  `LegacyDocumentFormat` y `LegacyConversionException`.
- `OfficeEngine.LibreOffice`, proyecto de infraestructura interna que ejecuta el motor con su perfil y
  sus temporales propios, compartido con el adaptador de PDF.
- Verificación diferencial: el motor escribe el archivo heredado y el conversor lo lee.

La decisión completa esta en `docs/adr/0042-legacy-format-conversión.md`.

## Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **952 superadas, 0 con error, 0 omitidas** (952 total). |

Las **945 pruebas anteriores siguen pasando**; las 7 nuevas están en `LegacyConverterTests`.

Salida textual del resumen (la duracion varia):

```
Correctas! - Con error:     0, Superado:   952, Omitido:     0, Total:   952
```

## Lo que se entrega, y lo que no

| Capacidad | Estado |
| --- | --- |
| Convertir `.doc` a `.docx` | Probado |
| Convertir `.xls` a `.xlsx` | Probado |
| Leer un `.doc` o un `.xls` directamente | **No implementado** |
| Editar un archivo heredado en su formato | **No implementado** |
| Guardar en el formato heredado | **No implementado** |

La distinción no es académica: leer, editar y guardar en el formato original son capacidades
distintas, y solo la primera existe en forma de conversión (ADR 0006 y `AGENT.md`).

## Verificación independiente (no autorreferencial)

1. **Viaje de ida y vuelta por el formato heredado**: el motor escribe el `.doc` y el `.xls` de
   partida, el conversor los lee, y el resultado se abre con `WordDocument.Open` y
   `ExcelWorkbook.Open` para comprobar que el texto y los valores siguen ahi. Ningún binario
   heredado se guarda en el repositorio: los fixtures los produce el motor en cada ejecucion.
2. **Validacion de esquema** del paquete convertido con `OpenXmlValidator`.
3. **El motor es tolerante**: un archivo que no es un `.doc` valido puede abrirlo como texto. La
   prueba comprueba que el conversor nunca devuelve algo que no sea un documento, admita el refus o
   la conversión tolerante.

## Hallazgos durante las pruebas y acciones tomadas

1. **El motor se duplicaba.** El adaptador de PDF ya tenia el aislamiento del proceso (perfil
   propio, temporales, tiempo de espera, código de salida). Se extrajo a `OfficeEngine.LibreOffice` y
   los dos adaptadores lo usan: hay una sola implementacion de como se arranca y se vigila el
   motor, y el cambio no rompió ninguna de las 17 pruebas de PDF.
2. **El archivo de salida no se busca por extensión.** El motor deja temporales (`lu*.tmp`) en el
   directorio de salida, así que el conversor espera el nombre exacto (`source.docx`, `source.xlsx`)
   en vez de accepting el primer `.docx` que encuentre.
3. **El motor acepta cualquier cosa.** Igual que con el PDF, unos bytes que no son un `.doc` se
   abren como texto y se convierten. La prueba se reescribio para comprobar lo que importa: que el
   conversor nunca devuelve algo que no sea un documento.

## Qué **no** se ha implementado (y por tanto no se afirma)

- Leer, editar y guardar en los formatos heredados.
- Los formatos heredados de PowerPoint, que ni siquiera se convierten.
- Una matriz de las pérdidas por formato: las pérdidas dependen del motor y del archivo, y se
  documentan de forma general (ADR 0042), no caso por caso.
- El **motor de fórmulas propio** (6.2), que es un subsistema aparte con su matriz de funciones
  públicada.


## Incremento 6.2 - Motor de fórmulas propio

### Alcance

Un subsistema de cálculo separado, con la matriz de funciones públicada y evidencia diferencial.
Incluye:

- `OfficeEngine.Fórmulas`: `FórmulaEngine`, `FórmulaEngineOptions`, `IFórmulaContext`, `FórmulaError`
  y `FórmulaSyntaxException`. El proyecto no depende de Word, ni de Excel, ni del motor externo.
- `OfficeEngine.Excel`: `ExcelFórmulaContext` y `ExcelWorkbook.EvaluateFórmula`, que evalúa sobre los
  valores del libro.
- `docs/formula-matrix.md`: la matriz públicada con lo que se soporta y lo que no.
- `FórmulaDifferentialTests`: 31 fórmulas comparadas con un motor independiente.

La decisión completa está en `docs/adr/0043-formula-engine.md`.

### Comandos y resultados

| Comando | Resultado |
| --- | --- |
| `dotnet build OfficeEngine.sln -c Release` | Compilación correcta. 0 advertencias, 0 errores. |
| `dotnet test OfficeEngine.sln -c Release --no-build` | **1034 superadas, 0 con error, 0 omitidas** (1034 total). |
| `dotnet test ... --collect:"XPlat Code Coverage" --results-directory artifacts/testresults/phase6/inc62` | **1034 superadas**; cobertura en `artifacts/testresults/phase6/inc62/**/coverage.cobertura.xml`. |

Las **952 pruebas anteriores siguen pasando**; las 82 nuevas estan en `FórmulaEngineTests` (76),
`ExcelFórmulaTests` (5) y `FórmulaDifferentialTests` (1).

### Cobertura (observada, XPlat Code Coverage)

| Archivo | Líneas cubiertas | % | Ramas |
| --- | --- | --- | --- |
| `Fórmulas/FórmulaEngine.cs` | 13 / 13 | 100 % | 100 % |
| `Fórmulas/FórmulaEvaluator.cs` | 211 / 254 | 83,1 % | 72,3 % |
| `Fórmulas/FórmulaFunctions.cs` | 158 / 184 | 85,9 % | 62,3 % |
| `Fórmulas/FórmulaTokenizer.cs` | 72 / 118 | 61,0 % | 62,9 % |
| `Fórmulas/FórmulaValues.cs` | 54 / 89 | 60,7 % | 42,3 % |
| **Total de la solución** | 6607 / 7168 | 92,2 % | 80,9 % |

Las líneas sin cubrir del motor son sobre todo rutas de error del analizador (una cadena sin cerrar,
un número inválido, un carácter inesperado) y las funciones que la matriz declara pero que ninguna
prueba ejercita todavía.

### Verificación independiente (no autorreferencial)

1. **Diferencial con un motor externo**: 31 fórmulas (aritmética, referencias, rangos, funciones,
   comparaciónes, concatenacion y errores) se escriben en un libro, el motor externo las recalcula al
   cargar y las exporta a CSV, y el valor del CSV se compara con el del motor de OfficeEngine **y**
   con el valor esperado escrito a mano. Los tres coinciden.
2. **Normalización del idioma**: el exportador escribe `VERDADERO` y `#¡DIV/0!` en espanol; la
   comparación los traduce a `TRUE` y `#DIV/0!`. Es una diferencia de idioma, no de significado, y
   queda anotada en la matriz.

### Hallazgos durante las pruebas y acciones tomadas

1. **Un rango no se expandía.** El parser consumía el identificador y los dos puntos antes de leer
   el rango, y solo leía la última celda: `SUM(A1:A2)` daba 5 en vez de 5 por el camino equivocado, y
   `COUNT(A1:A2)` contaba 1. Corregido: el rango se lee desde el primer identificador.
2. **Los errores se escapaban antes de que `ISERROR` los viera.** Los errores se lanzaban como
   excepción, así que `ISERROR(1/0)` devolvía `#DIV/0!` en vez de `TRUE`. Ahora el error es un
   **valor** que viaja por la expresión, y las funciones que lo aceptan (`ISERROR`, `IFERROR`) lo
   reciben.
3. **`IF` evalúaba las dos ramas.** `IF(A1>0;1;1/0)` devolvía `#DIV/0!` en vez de `1`. Ahora las
   ramas se evalúan de forma perezosa: se localiza donde acaba cada argumento y solo se evalúa el
   necesario.
4. **Los nombres en español de la matriz no se aceptaban** (hallazgo posterior, al documentar la
   biblioteca). La matriz y el ADR 0043 declaraban que el motor aceptaba el nombre de cada función
   en inglés y en español, pero solo se aceptaba el inglés: `=SUMA(A1:A2)` devolvía `#NAME?`. Como
   Excel guarda los nombres en inglés dentro del archivo, escribir la fórmula tal cual además dejaba
   un libro que Excel no podía calcular. Ahora el motor acepta los dos nombres y, al escribir la
   fórmula en un libro, la traduce al nombre canónico; el texto de los literales no se toca. La
   matriz queda como estaba, y ahora es cierta.

### Qué **no** se ha implementado (y por tanto no se afirma)

- Funciones volátiles (`HOY`, `AHORA`, `ALEATORIO`): sus resultados no serían reproducibles.
- Matrices dinámicas, referencias externas, nombres definidos y fórmulas de matriz.
- Seguimiento de dependencias: el motor evalúa sobre los valores de las celdas, no sobre las fórmulas
  que contienen, y no recalcula nada.
- Fechas: el motor no conoce el sistema de fechas del libro, así que una celda con fecha se lee como
  número.

## Limitaciones conocidas

- La conversión necesita el motor instalado: sin él, `LegacyConverter.Convert` lanza
  `LegacyConversionException`.
- Cada conversión arranca un proceso del motor: cuesta unos segundos.
- El resultado es la interpretación del motor, no una equivalencia con el original: las macros, los
  objetos incrustados, los vínculos OLE, la maquetación exacta y los metadatos se pierden.
