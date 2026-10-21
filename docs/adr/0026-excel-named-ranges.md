# ADR 0026 — Rangos con nombre

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.6a

## Contexto

Un libro puede nombrar un rango (`Datos` → `Hoja1!$A$1:$B$2`) para poder referirse a él sin
escribir la referencia. En SpreadsheetML los nombres viven en `xl/workbook.xml`, dentro de
`<definedNames>`, y no en la hoja.

El elemento `<definedName>` es más amplio que «un rango con nombre»: tambiénalmacena constantes
(`42`), fórmulas (`=Hoja1!$A$1+1`), referencias a varias áreas, referencias a otros libros y los
nombres integrados que usa Excel (`_xlnm.Print_Area`, `_xlnm.Print_Titles`). Además, un nombre puede
tener **ámbito de libro** o **ámbito de hoja**: en el caso de la hoja, el atributo `localSheetId` no
apunta a un identificador estable, sino a la **posición de la hoja en la lista**, que cambia cuando
se elimina una hoja anterior.

Modelar solo «rangos con nombre» sin decidir qué pasa con el resto de nombres dejaría archivos
ajenos con definiciones que se pierden o que se interpretan mal.

## Decisión

### 1. API

- `ExcelWorkbook.DefinedNames`: catálogo de solo lectura con los nombres del libro, en el orden en
  que están almacenados.
- `ExcelWorkbook.FindDefinedName(string name)`: busca un nombre por su texto, sin distinguir
  mayúsculas; devuelve `null` si no existe.
- `ExcelWorkbook.RemoveNamedRange(string name)`: elimina un nombre de ámbito de libro.
- `Worksheet.AddNamedRange(string name, string rangeReference)`: crea un nombre de ámbito de libro
  sobre un rango de esa hoja.
- `Worksheet.AddLocalNamedRange(string name, string rangeReference)`: crea un nombre de ámbito de
  esa hoja.
- `Worksheet.RemoveLocalNamedRange(string name)`: elimina un nombre de ámbito de esa hoja.

`ExcelDefinedName` expone:

| Propiedad | Significado |
| --- | --- |
| `Name` | El texto del nombre, por ejemplo `Datos`. |
| `RefersTo` | El texto tal como se guarda, por ejemplo `Sheet1!$A$1:$B$2`. |
| `Scope` | La hoja del ámbito, o `null` si el nombre es de ámbito de libro. |
| `IsRange` | Si el nombre apunta a una celda o a un rango contiguo de una hoja de este libro. |
| `Range` | El rango en forma A1 relativa (`A1:B2`), o `null` si no es un rango. |

### 2. El texto del nombre se valida antes de tocar el paquete

- No vacío, como máximo 255 caracteres y con caracteres válidos en XML.
- Solo letras, dígitos, espacios, guiones bajos y puntos; debe empezar por una letra, un guion bajo
  o un punto.
- No puede parecer una referencia de celda (`A1`, `AB12`): Excel lo rechaza y la librería también.

### 3. El destino se guarda en forma absoluta

- El rango se da en forma A1 relativa (`A1:B2`), con las mismas reglas que `Worksheet.GetRange`, y
  también se admite una sola celda (`C3`).
- Se guarda como Excel lo guarda: `$A$1:$B$2`, con la hoja delante para el ámbito de libro
  (`Sheet1!$A$1:$B$2`) y sin ella para el ámbito de hoja (`$A$1:$B$2`).
- El nombre de la hoja se entrecomilla cuando lo necesita: `'Mi Hoja'!$A$1`.

### 4. No puede haber nombres duplicados

Asignar un nombre que ya existe **en el mismo ámbito** reemplaza su destino; no crea un duplicado ni
lanza un error. Un nombre de libro y un nombre de hoja pueden coincidir en el texto: son nombres
distintos.

### 5. El ámbito de hoja es posicional y se ajusta al eliminar hojas

`localSheetId` es la posición de la hoja en la lista. Por eso, al eliminar una hoja:

- se eliminan los nombres que eran de esa hoja, y
- se decrementa el ámbito de los nombres de las hojas posteriores, para que sigan apuntando a la
  misma hoja.

### 6. Lo que no se modela se conserva

El catálogo incluye **todos** los nombres, incluidos los que esta versión no interpreta: constantes,
fórmulas, referencias a varias áreas, referencias a otros libros, nombres ocultos y nombres
integrados como `_xlnm.Print_Area`. Para esos, `IsRange` es `false` y `Range` es `null`; su texto se
conserva tal cual al guardar.

No se interpretan ni se modifican: no hay referencias 3D, ni referencias externas, ni nombres
dinámicos, ni fórmulas con nombre.

## Consecuencias

- La API pública nueva es `ExcelDefinedName`, `ExcelWorkbook.DefinedNames`,
  `ExcelWorkbook.FindDefinedName`, `ExcelWorkbook.RemoveNamedRange`, `Worksheet.AddNamedRange`,
  `Worksheet.AddLocalNamedRange` y `Worksheet.RemoveLocalNamedRange`.
- El `xl/workbook.xml` producido puede contener `<definedNames>`, que se inserta en el orden que
  exige el esquema.
- Los nombres que no son rangos se listan pero no se interpretan; modificarlos o crearlos exige el
  API de nombres, no una edición del XML.
- Reordenar hojas no está implementado, así que el ámbito posicional solo cambia al eliminar una
  hoja, y ese caso queda cubierto.

## Verificación

- `ExcelDefinedNameTests`: creación de nombre de libro y de hoja; celda única; reemplazo del
  destino con el mismo nombre; comillas del nombre de hoja; validación del texto del nombre y del
  rango sin modificar el libro; eliminación de nombres de libro y de hoja; lectura de los nombres de
  un archivo ajeno (rango, constante y nombre integrado); preservación al guardar; ajuste del
  ámbito al eliminar una hoja; libro liberado.
- `OoxmlValidation.CreateXlsxWithDefinedNames` construye el fixture con el SDK y
  `ReadXlsxDefinedNames` lo lee con el SDK, no con la librería bajo prueba.
