# ADR 0042 — Conversión de los formatos heredados con un motor externo

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 6, incremento 6.1

## Contexto

Word 97-2003 (`.doc`) y Excel 97-2003 (`.xls`) son archivos OLE2: un contenedor compuesto con
flujos binarios, no un paquete XML. El SDK de Open XML no los lee, y la librería no tiene un
lector de OLE2.

El `AGENT.md` pide separar con claridad cuatro capacidades que no son equivalentes: **convertir**,
**leer**, **editar** y **guardar en el formato original**. Y el `PROMPT.md` sitúa el soporte de
DOC/XLS en una fase posterior, con pérdidas documentadas.

Con el motor externo ya verificado (ADR 0005, ADR 0038), la conversión es posible. La lectura
directa del formato heredado, no.

## Decisión

### 1. Solo se convierte

`OfficeEngine.Legacy` ofrece **una** operación: `LegacyConverter.Convert(Stream source,
LegacyDocumentFormat format)`, que convierte un `.doc` en un `.docx` o un `.xls` en un `.xlsx`.

- El resultado es un **documento nuevo**, la interpretación que hace el motor del original. No es una
  traducción byte a byte.
- El formato heredado **no se edita ni se guarda**: para trabajar con el contenido se convierte
  primero y luego se abre el resultado con `WordDocument.Open` o `ExcelWorkbook.Open`. Guardar
  escribe el formato moderno.

Quedan fuera, y no se afirman: leer un `.doc` directamente, editarlo en su formato y guardarlo
así.

### 2. El motor se comparte, no se duplica

El aislamiento del proceso (directorio temporal propio, perfil de usuario propio, tiempo de espera,
código de salida, diagnóstico) ya existía en el adaptador de PDF. Se extrae a
`OfficeEngine.LibreOffice`, un proyecto de infraestructura interna, y los dos adaptadores lo usan.
Hay **una** implementación de cómo se arranca y se vigila el motor.

### 3. Las pérdidas se documentan, no se esconden

Lo que la conversión **no** conserva, porque depende del motor y del archivo: las macros, los
objetos incrustados y los vínculos OLE, la maquetación exacta, los metadatos del original y
cualquier rasgo que el motor no sepa expresar. La conversión es un **esfuerzo del motor**, no una
garantía.

### 4. La verificación es un viaje de ida y vuelta por el formato heredado

Las pruebas no se apoyan en un archivo binario guardado en el repositorio: el propio motor escribe
el `.doc` o el `.xls` de partida, el conversor lo lee, y el resultado se abre con la API moderna
para comprobar que el texto y los valores siguen ahí. La evidencia es, por tanto, **diferencial**:
dos pasos del mismo motor, uno independiente de la librería.

## Consecuencias

- La API pública nueva es `LegacyConverter`, `LegacyConverterOptions`, `LegacyConversionResult`,
  `LegacyDocumentFormat` y `LegacyConversionException`.
- El paquete `OfficeEngine.Legacy` necesita el motor instalado; sin él, la conversión falla con
  `LegacyConversionException`.
- Un documento heredado con contenido que el motor no conserva llega convertido y **sin** ese
  contenido: es la limitación conocida del formato, no un fallo de la librería.

## Verificación

- `LegacyConverterTests`: un documento de Word y un libro de Excel viajan a su formato heredado y
  vuelven, y el resultado se abre con la API moderna conservando texto y valores; un archivo
  corrupto nunca devuelve algo que no sea un documento; y los errores de la API (flujo vacío, nulo,
  formato desconocido, motor ausente) se rechazan con su tipo.
- `LegacyFixtures` produce los archivos heredados con el motor, sin guardar binarios en el repositorio.
