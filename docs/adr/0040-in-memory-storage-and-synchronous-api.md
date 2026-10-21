# ADR 0040 — Modelo de almacenamiento en memoria y API síncrona

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 5, incremento 5.3

## Contexto

Desde el ADR 0004 quedó abierta una decisión (la número 2 de `docs/risks-and-dependencies.md`): qué
pasa con los documentos grandes y si la API debe ser asíncrona.

La librería mantiene el paquete **íntegro en memoria** durante toda la vida del documento. Eso tiene
un coste medible: el ZIP descomprimido, el DOM del SDK y el buffer propio conviven, y la memoria
usada es un múltiplo del tamaño del archivo.

La pregunta no es solo de rendimiento, sino de contrato: si la API fuese asíncrona, cada método
público cambiaría de firma.

## Decisión

### 1. El modelo en memoria se mantiene, y es una decisión consciente

El paquete es la unidad de trabajo: se abre, se edita y se guarda. Un almacén respaldado por disco
o por segmentos cambiaría la semántica del guardado (qué se ve desde fuera mientras se edita), el
ciclo de vida y el manejo de errores, a cambio de un beneficio que solo aparece con documentos más
grandes que la memoria disponible.

El ADR 0004 ya lo dejó anotado como trabajo futuro, y se mantiene como tal: el día que el almacén
cambie, se añadirán las variantes `...Async` **sin romper** la firma síncrona actual.

### 2. El límite que manda es el del tamaño expandido

El límite real de memoria no es `MaxFileSizeBytes` (el tamaño del archivo en disco) sino
`MaxTotalUncompressedSizeBytes` (el tamaño al que se expande el ZIP), y por encima de él está el DOM
del SDK, que multiplica el tamaño del XML. Los tres por omisión (256 MiB de archivo, 512 MiB
expandidos, 4096 entradas) son una **política**, no una promesa de memoria: un paquete en el límite
puede ocupar bastante más de 512 MiB.

### 3. La API sigue siendo síncrona

No se publica ningún método `async`: el trabajo es de memoria, no de espera de E/S, y una promesa
asíncrona sin espera real sería engañosa (ADR 0004, punto 1).

### 4. Los costes se miden y se publican

El proyecto `OfficeEngine.Benchmarks` mide los caminos críticos y su resultado está en
`benchmarks/baseline.txt` y en `docs/phase-5-report.md`. Los números se miden **en el entorno de
desarrollo** (Windows 11, .NET 8) y no son una promesa para otras máquinas.

## Consecuencias

- Positivas: contrato estable, sin semántica de guardado surprising, y límites comprobados con
  pruebas (`LargeDocumentTests`).
- Negativas: un documento grande ocupa varias veces su tamaño en memoria, y abrir uno es lineal en
  su tamaño; no hay cancelación de una apertura en curso.
- La escritura de **muchas** celdas sigue siendo celda a celda (ADR 0022) y, por tanto, superlineal:
  el benchmark lo mide y el ADR 0022 lo declara como atajo conocido.

## Verificación

- `LargeDocumentTests`: un documento de Word de 10 000 párrafos y un libro de Excel de 10 000 filas
  abren, se guardan y se vuelven a abrir sin pérdida; y un paquete que se expande por encima del
  límite se rechaza.
- `ExcelPerformanceTests`: escribir 3 000 celdas no asigna memoria de forma cuadrática.
- `OfficeEngine.Benchmarks`: las seis escenarios medidas, con su referencia.
