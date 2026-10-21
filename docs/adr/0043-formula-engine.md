# ADR 0043 — Motor de fórmulas: el subconjunto soportado y su matriz

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 6, incremento 6.2

## Contexto

La Fase 3 escribe fórmulas en las celdas pero no las calcula, y el `PROMPT.md` exige que el cálculo
sea un subsistema aparte, con una **matriz de funciones publicada** y pruebas reproducibles. Con el
motor externo verificado (ADR 0038) se puede, además, contrastar los resultados con un motor
independiente: la evidencia diferencial que el `AGENT.md` pide.

El motor de fórmulas de Excel es enorme (más de 500 funciones, matrices dinámicas, referencias
externas, nombres definidos, funciones volátiles). Declarar que se calcula "como Excel" sin pruebas
sería falso.

## Decisión

### 1. Un subsistema aparte, que no sabe de Word ni de Excel

`OfficeEngine.Formulas` contiene el motor y **nada más**: no depende de `OfficeEngine.Core`, ni de
Word, ni de Excel. Los valores se los pide a un `IFormulaContext`, que el llamante implementa.

- `OfficeEngine.Excel` aporta `ExcelFormulaContext` y el atajo `ExcelWorkbook.EvaluateFormula`, que
  evalúa sobre los **valores** del libro.
- El motor no sigue dependencias: no calcula la fórmula de una celda para usarla de argumento,
  evalúa sobre lo que hay. Una referencia circular no es un problema, y nada se recalcula.

### 2. Un subconjunto explícito, con la matriz publicada

`docs/formula-matrix.md` publica exactamente lo que se soporta: los literales, las referencias, los
operadores, las 32 funciones (con el nombre en inglés y en español: 61 nombres) y las reglas de
conversión. Lo que no está en la matriz no se calcula, y no se afirma.

Los nombres se aceptan en inglés y en español porque el público de la biblioteca es hispanohablante
y las hojas se escriben en ese idioma. Como Excel guarda los nombres de las funciones en inglés,
la biblioteca **traduce** la fórmula al nombre canónico cuando la escribe en un libro (el texto de
los literales no se toca).

### 3. Los errores son valores, no excepciones

Un error de Excel (`#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`) es un
**valor** que viaja por la expresión. Por eso `ISERROR` y `IFERROR` pueden verlo, y por eso el
resultado de una fórmula con un error es ese error, no una excepción.

### 4. `IF` no evalúa la rama que no se toma

`=SI(A1>0;1;1/0)` devuelve `1`, no `#DIV/0!`. Para eso, `IF` e `IFERROR` se evalúan de forma
perezosa: se localiza dónde acaba cada argumento y solo se evalúa el necesario.

### 5. La evidencia es diferencial

`FormulaDifferentialTests` escribe un libro con 31 fórmulas, lo exporta a CSV con **el motor
externo** (que recalcula al cargar) y compara el valor del CSV con el del motor de OfficeEngine, y
con el esperado escrito a mano.

La comparación normaliza el **idioma** del exportador: este escribe `VERDADERO` y `#¡DIV/0!` en
español, y se traducen a `TRUE` y `#DIV/0!` antes de comparar. Es una diferencia de idioma, no de
significado.

## Consecuencias

- La API pública nueva es `FormulaEngine`, `FormulaEngineOptions`, `IFormulaContext`,
  `FormulaError`, `FormulaSyntaxException`, `ExcelFormulaContext` y `ExcelWorkbook.EvaluateFormula`.
- Un libro con fórmulas se puede calcular sin escribir fórmulas: el motor lee los valores.
- La matriz es un contrato: cambiarla es un cambio de alcance, no un detalle interno.
- Lo que no está en la matriz (volátiles, matrices, referencias externas, nombres definidos) devuelve
  `#NAME?` o `#VALUE!`, nunca un valor inventado.

## Verificación

- `FormulaEngineTests`: 58 casos con los valores, los operadores, los errores de Excel, la
  evaluación perezosa de `IF` y los límites del motor, más un caso por función que comprueba que el
  nombre en español da el mismo resultado que el nombre en inglés.
- `ExcelFormulaTests`: el motor leyendo los valores de un libro, con hoja explícita y con celdas
  vacías, y la traducción del nombre de la función al escribir la fórmula.
- `FormulaNameTests`: la traducción de los nombres al nombre canónico, sin tocar los literales.
- `FormulaDifferentialTests`: 31 fórmulas comparadas contra un motor independiente y contra el
  valor esperado.
