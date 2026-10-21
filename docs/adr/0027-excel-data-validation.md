# ADR 0027 — Validación de datos

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.6b

## Contexto

Una hoja puede exigir que las celdas de un rango cumplan una regla: una lista de valores permitidos,
un número entero dentro de unos límites, una fecha, un texto de una longitud dada o una fórmula. En
SpreadsheetML las reglas viven en la hoja, dentro de `<dataValidations>`, y cada `<dataValidation>`
aplica a un conjunto de rangos (`sqref`).

Cada regla tiene además mensajes para el usuario (título y texto de error, título y texto de
entrada) y opciones de presentación (`allowBlank`, `showInputMessage`, `showErrorMessage`). Modelar
todo eso de golpe es mucho superficie para un uso mayoritario muy concreto: la lista de valores.

## Decisión

### 1. API

- `ExcelDataValidation.List(rangeReference, values, allowBlank = true)`.
- `ExcelDataValidation.WholeNumber(rangeReference, minimum = null, maximum = null, allowBlank = true)`.
- `ExcelDataValidation.DecimalNumber(rangeReference, minimum = null, maximum = null, allowBlank = true)`.
- `Worksheet.AddDataValidation(validation)`: añade la regla a la hoja.
- `Worksheet.DataValidations`: catálogo de solo lectura con las reglas de la hoja.
- `Worksheet.RemoveDataValidation(rangeReference)`: elimina la regla de un rango.

`ExcelDataValidation` es inmutable y solo se crea con las fábricas de arriba. Expone `Range`, `Type`,
`Formula1`, `Formula2`, `Minimum`, `Maximum`, `AllowedValues`, `AllowBlank`, `ErrorTitle`, `Error`,
`PromptTitle` y `Prompt`. Los mensajes se conservan al leer, pero esta versión no los escribe: las
reglas que crea no llevan mensaje.

### 2. El operador se deduce de los límites

No hay un `Operator` en la API. El significado lo dan `Minimum` y `Maximum`:

| Mínimo | Máximo | Operador escrito | Fórmula 1 | Fórmula 2 |
| --- | --- | --- | --- | --- |
| sí | sí | `between` | mínimo | máximo |
| sí | no | `greaterThanOrEqual` | mínimo | — |
| no | sí | `lessThanOrEqual` | máximo | — |
| no | no | (ninguno) | — | — |

Con los dos límites, `between` es el operador de Excel. Con uno solo, el operador «mayor o igual» o
«menor o igual» es el que corresponde y evita una regla que Excel interpretaría de otra forma.

### 3. La lista se guarda en línea

Una lista se guarda como una constante entre comillas con los valores separados por comas:
`"uno,dos,tres"`. Por eso un valor **no puede contener una coma ni una comilla doble**, y la API lo
rechaza con `ArgumentException` en vez de escribir una lista que Excel no interpretaría igual.

### 4. El rango se valida antes de tocar el paquete

El rango se da en forma A1 (`A1:A10`) y se valida con las mismas reglas que `Worksheet.GetRange`; una
sola celda (`C3`) también vale. Un rango inválido, una lista vacía, un valor no admitible o un mínimo
mayor que el máximo lanzan `ArgumentException` y no modifican la hoja.

### 5. Lo que no se modela se conserva

- Las reglas de tipo `date`, `time`, `textLength` y `custom` se **leen** con su tipo, sus fórmulas y
  sus mensajes, pero no se interpretan: no hay forma de crearlas ni de modificarlas.
- Una regla que aplica a varias áreas (`sqref="A1:A10 C1:C10"`) informa de la primera; el resto de
  las áreas sigue en el archivo.
- Los mensajes y las opciones de presentación de una regla ajena se conservan al guardar, pero no se
  exponen como propiedades editables más allá de los textos.
- La validación no impide escribir: Excel la aplica al escribir, no al guardar. Esta librería no
  comprueba los valores de las celdas ni bloquea la edición.

## Consecuencias

- La API pública nueva es `ExcelDataValidation`, `ExcelDataValidationType`,
  `Worksheet.AddDataValidation`, `Worksheet.DataValidations` y `Worksheet.RemoveDataValidation`.
- La hoja producida puede contener `<dataValidations>`, que se inserta en el orden que exige el
  esquema y ya no se reporta como contenido no modelado.
- Las reglas que esta versión no modela se leen y se conservan, pero no se pueden crear.

## Verificación

- `ExcelDataValidationTests`: lista con round-trip y valores leídos; número entero con los tres
  operadores; regla decimal; rango inválido; lista inválida (vacía, con coma, con comilla doble, con
  valor vacío); mínimo mayor que el máximo; lectura de las reglas de un archivo ajeno, incluidas una
  de tipo `custom`; preservación al guardar; eliminación de una regla; ausencia del diagnóstico de
  contenido no modelado; libro liberado.
- `OoxmlValidation.CreateXlsxWithDataValidations` construye el fixture con el SDK y
  `ReadXlsxDataValidations` lee tipo, `sqref`, fórmulas y operador con el SDK.
