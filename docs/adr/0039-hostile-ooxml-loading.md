# ADR 0039 — Carga de OOXML hostil y corpus de archivos corruptos

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 5, incrementos 5.1 y 5.2

## Contexto

OfficeEngine abre archivos que no controla. El ADR 0007/settings fijaron límites de tamaño, de
entradas y de ratio de compresión, pero no cubrieron dos riesgos que el sondeo empírico del
2026-10-21 destapó:

1. **El SDK carga las partes con un lector recursivo.** Un documento con unos miles de elementos
   anidados agota la pila y **termina el proceso** con un `StackOverflowException` (0xC00000FD), que
   no se puede capturar. Medido en este entorno: 1 000 niveles se abren bien; 2 000 matan el proceso.
   Un archivo hostil es, por tanto, una denegación de servicio, no un error.
2. **El SDK lanza excepciones propias** (`InvalidDataException`, `InvalidOperationException`,
   `ArgumentOutOfRangeException`, `XmlException`) que se escapaban de la API pública. La promesa de
   la librería es de errores tipados, y un `XmlException` crudo la incumple.

El SDK ya prohíbe los DTD (`DtdProcessing.Prohibit`), así que *billion laughs* y entidades externas
no se expanden: fallan. Pero fallan con un `XmlException` crudo, y el momento del fallo (al
materializar el DOM de la parte) es perezoso: podía escaparse mucho después de abrir el archivo.

## Decisión

### 1. Las partes XML se miden antes de que las parsee el lector

`OoxmlXmlGuard` (en `OfficeEngine.OpenXml`) recorre el paquete **antes** de abrirlo y mide cada
parte que parece XML con un `XmlReader` de flujo, que no usa la pila:

- **Sin DTD**: `DtdProcessing.Prohibit` y `XmlResolver = null`. Una declaración de tipo de documento
  se rechaza con `InvalidOfficeFileException` y un mensaje que nombra la parte.
- **Con profundidad acotada**: si una parte anida más de `OfficeEngineLimits.MaximumXmlDepth` niveles
  (256 por omisión), se rechaza con `OfficeResourceLimitExceededException`.

Una parte se considera XML cuando su primer byte no blanco es `<`, así que una parte **renombrada**
para ocultar su tipo también se mide. Las partes binarias (imágenes, medios) se saltan.

El coste es una pasada de flujo extra por las partes XML. La suite completa tarda lo mismo que antes
(20 s), porque el SDK vuelve a parsear esas partes al materializar su DOM.

El límite por omisión es 256 porque un documento real de Word o Excel anida menos de 30 niveles, y
el proceso muere a partir de unos 1 500: el margen es de más de cinco veces.

### 2. El contrato de errores tipados se cumple en las fronteras

- `OoxmlPartReader.Read` envuelve la materialización del DOM de una parte y traduce
  `InvalidDataException`, `XmlException`, `OpenXmlPackageException`, `FileFormatException` y
  `ArgumentOutOfRangeException` a `InvalidOfficeFileException`, conservando la excepción original
  como `InnerException`.
- La apertura (`WordDocument.Open`, `ExcelWorkbook.Open`) traduce las mismas familias, y la
  inspección inicial del paquete también.
- El mensaje público **no** incluye el texto del documento: el detalle del error de XML queda en la
  excepción interna, que es donde el llamante puede pedirlo de forma explícita.

### 3. Un corpus determinista de archivos corruptos

`CorruptPackageTests` muta un `.docx` y un `.xlsx` válidos de forma **reproducible** (sin
aleatoriedad): truncados en nueve puntos, un byte alterado en diez puntos, tres entradas
eliminadas y tres entradas sustituidas por contenido que no es lo que declaran.

El contrato es: cada mutación **o falla con un error tipado de OfficeEngine, o se abre**. Nunca hay
una excepción del SDK o de la biblioteca de XML, y nunca hay un fallo del proceso.

El corpus encontró tres fugas reales (`InvalidDataException` e `InvalidOperationException` del SDK
en la apertura y en la inspección) que ahora están traducidas.

### 4. El límite es público y configurable

`OfficeEngineLimits.MaximumXmlDepth` (256 por omisión) es parte del contrato público, como los demás
límites, y se valida con `EnsureValid`.

## Consecuencias

- Un archivo con XML anidado en exceso **falla con un error tipado** en lugar de tumbar el proceso.
- Un DTD se rechaza de forma explícita, con el nombre de la parte, en lugar de depender del
  comportamiento del SDK.
- La API ya no filtra excepciones del SDK ni del lector de XML.
- Abrir un paquete cuesta una pasada de flujo más por sus partes XML.
- Un límite de profundidad demasiado bajo rechazaría documentos válidos: el valor por omisión está
  muy por encima de lo que usa Word o Excel.

## Verificación

- `HostilePackageTests`: anidamiento profundo en un `.docx` y en un `.xlsx`; declaración de tipo de
  documento; entidad externa; carácter de control inválido; límite configurable; y un documento
  válido que se sigue abriendo.
- `CorruptPackageTests`: el corpus de mutaciones de un `.docx` y de un `.xlsx`, y la comprobación de
  que el corpus tiene el tamaño esperado.
- `OfficeEngineLimitsTests`: el límite de profundidad por omisión y su validación.
