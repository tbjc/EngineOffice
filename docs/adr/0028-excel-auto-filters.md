# ADR 0028 — Autofiltros

- Estado: Aceptado
- Fecha: 2026-10-20
- Fase: 3, incremento 3.6c

## Contexto

Una hoja puede tener un autofiltro: los botones de la cabecera de un rango y el filtro que el usuario
aplica sobre él. En SpreadsheetML es un elemento `<autoFilter>` dentro de la hoja, con el atributo
`ref` que delimita el rango.

El elemento también guarda los **criterios** que el usuario aplicó (`<filterColumn>` con filtros o
condiciones) y, como hermano, el estado de ordenación (`<sortState>`). Los criterios son estado de
la sesión de filtrado, no estructura de la hoja: dependen de lo que el usuario escribió.

Modelar los criterios exigiría decidir cómo se represents una lista de valores filtrados, una
condición o un filtro de color, y none de eso tiene valor para un archivo que se genera o se
transforma.

## Decisión

### 1. API

- `ExcelAutoFilter`: una clase con `Range`, el rango que cubre el filtro, y un constructor que
  valida la referencia A1.
- `Worksheet.AutoFilter`: propiedad que devuelve el filtro o `null` cuando la hoja no tiene, y que
  permite asignar uno nuevo o `null` para quitarlo.

### 2. Solo se modela el rango

- El rango se da en forma A1 (`A1:D10`) y se valida como el resto de referencias. Que incluya la
  fila de cabecera es una condición de Excel, no se comprueba.
- Los **criterios** de filtrado y el estado de ordenación **no se modelan**: se conservan tal cual
  cuando la hoja se guarda sin tocarla y **se descartan** cuando el filtro se escribe otra vez, porque
  un criterio pertenece a un rango que puede haber cambiado.

### 3. El elemento se coloca en su sitio

`<autoFilter>` va antes de `mergeCells` en la hoja. Se inserta respetando el orden del esquema, y a
partir de este incremento la hoja ya no lo reporta como contenido no modelado.

## Consecuencias

- La API pública nueva es `ExcelAutoFilter` y `Worksheet.AutoFilter`.
- La hoja producida puede contener `<autoFilter>`; un archivo con filtros ajenos conserva los
  criterios al guardarse.
- No hay forma de aplicar criterios de filtrado desde la API.

## Verificación

- `ExcelAutoFilterTests`: escritura del rango y round-trip; eliminación con `null`; lectura del
  filtro de un archivo ajeno; rango inválido; ausencia del diagnóstico de contenido no modelado;
  libro liberado.
- `OoxmlValidation.CreateXlsxWithAutoFilter` construye el fixture con el SDK y `ReadXlsxAutoFilter`
  lee el rango con el SDK.
