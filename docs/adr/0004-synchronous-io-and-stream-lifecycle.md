# ADR 0004 — I/O síncrona y ciclo de vida de streams

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 0 (aplicable a la Fase 1)

## Contexto

`PROMPT.md` pide decidir si las operaciones con I/O son síncronas o asíncronas y mantener la
política de forma coherente. En esta etapa el paquete se mantiene íntegro en memoria, así que
no hay una espera de I/O real que justifique `async`.

## Decisión

1. **Toda la entrada/salida es síncrona.** No se publican métodos `async` que internamente
   hagan trabajo en memoria: serían una promesa vacía.
2. `Open(path)` lee el archivo a un buffer propio. No se retiene un bloqueo de archivo, de modo
   que guardar sobre el mismo archivo es seguro.
3. `Open(stream, leaveOpen = false)` copia el stream y, por defecto, lo cierra. El documento
   **nunca** conserva una referencia al stream del llamador.
4. `Open(stream)` escribe el `Position` del stream de entrada a su estado final habitual
   (lectura completa) y luego lo cierra; no se intenta restaurar la posición.
5. `Save(path)` escribe en un temporal del mismo directorio y hace `File.Move(overwrite: true)`
   para que el destino no quede truncado si la escritura falla.
6. `Save(stream)` deja el stream destino abierto y solo escribe.
7. `Dispose` libera el paquete y el buffer. Un documento dispuesto lanza
   `ObjectDisposedException` en cualquier operación posterior.

## Consecuencias

- Positivas: semántica simple y predecible; sin `async` engañoso; sin bloqueos de archivo;
  guardado atómico.
- Negativas: el uso de memoria es del orden del tamaño del archivo; los archivos muy grandes
  consumen mucha memoria y están acotados por `OfficeEngineLimits.MaxFileSizeBytes`.
- Trabajo futuro: almacén respaldado por archivo/segmentos y, **entonces**, una API `async`
  real. La firma síncrona actual no se romperá: se añadirán variantes `...Async`.
