# ADR 0041 — Medición de rendimiento reproducible y regresión de round-trip

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 5, incremento 5.4

## Contexto

El `AGENT.md` pone el «rendimiento medido, no supuesto» por delante de la amplitud de
funcionalidades, y la hoja de ruta de la Fase 5 pide «benchmarks reproducibles» y una «regresión de
round-trip sobre documentos generados».

Hasta ahora el rendimiento era una afirmación: el informe de la Fase 2 decía que el acceso masivo es
superlineal (ADR 0022) sin medir cuánto cuesta. Con el proyecto de PDF y el de benchmarking, medir
deja de ser opcional.

## Decisión

### 1. Un proyecto de medición, sin dependencias nuevas

`OfficeEngine.Benchmarks` es una aplicación de consola que mide seis escenarios: crear, abrir y
recorrer un documento de Word de 5 000 párrafos, y lo mismo con un libro de Excel de 5 000 filas.

- **No** se añade BenchmarkDotNet ni ninguna otra dependencia: el proyecto solo usa la propia
  biblioteca y la BCL. La regla 9 del `AGENT.md` obliga a justificar cada dependencia, y un proyecto
  de medición no es motivo suficiente para añadir una.
- La medida es el **tiempo transcurrido** y la **memoria asignada** (`GC.GetTotalAllocatedBytes`), que
  es donde se ven los problemas de memoria; el tamaño del paquete se informa como contexto.
- La salida es una tabla legible y un archivo de referencia (`benchmarks/baseline.txt`) con formato
  `escenario|ms|asignado|paquete`.

### 2. La comparación es una herramienta, no una puerta

`--baseline <archivo>` compara con la referencia y avisa de las regresiones que superan el umbral
(2×), y devuelve un código de salida distinto de cero. Pero **CI no falla por ello**: el rendimiento
de un *runner* compartido varía demasiado para que una variación de tiempo tumbe una compilación. La
integración continua **publica el informe** como artefacto, para poder comparar un cambio con el
anterior.

### 3. Una red de regresión sobre documentos generados

- `DocumentGenerator` produce documentos con una secuencia pseudoaleatoria **determinista**
  (xorshift): párrafos, runs con formato, tablas, secciones. `WordRoundTripRegressionTests` los
  guarda, los reabre y compara el texto y la estructura; `ExcelRoundTripRegressionTests` hace lo
  mismo con valores, estilos y fórmulas.
- `ExcelPerformanceTests` vigila que escribir muchas celdas no tenga un coste cuadrático. Vive en
  una colección sin paralelismo para que la medida de memoria le pertenezca.

## Consecuencias

- El rendimiento deja de ser una afirmación: hay números, un archivo de referencia y una forma de
  compararlos.
- El primer uso del proyecto de medición **encontró un fallo real**: la búsqueda de una cadena en la
  tabla de cadenas compartidas recorría la tabla entera, y el texto de cada entrada se reconstruía en
  cada comparación. Escribir 10 000 celdas asignaba 1 745 MB. Con un índice, son 29,8 MB: 58 veces
  menos, y el tiempo baja de 2,42 s a 1,79 s.
- La regresión está protegida por una prueba que falla si el coste vuelve a ser cuadrático.

## Verificación

- `OfficeEngine.Benchmarks` con `--baseline benchmarks/baseline.txt` devuelve 0 con la referencia
  actual y un código distinto de cero si una escena se degrada más de 2×.
- `ExcelPerformanceTests`: escribir 3 000 celdas no supera los 250 MB asignados.
- `WordRoundTripRegressionTests` y `ExcelRoundTripRegressionTests`: doce documentos generados y un
  libro generado se guardan y se reabren sin perder texto, estructura, valores, estilos ni fórmulas.
