# ADR 0007 — Seguridad ante archivos no confiables

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 0 (aplicado en la Fase 1)

## Contexto

Un archivo Office es una entrada no confiable: puede ser enorme, una bomba ZIP, contener rutas
maliciosas, XML con entidades externas, macros u objetos incrustados. El proceso debe fallar de
forma segura y acotada.

## Decisión

### Frontera y límites (implementado y probado en Fase 1)

- Toda entrada pasa por `OoxmlLoader`, que la copia a un buffer acotado por
  `OfficeEngineLimits.MaxFileSizeBytes`. El límite se comprueba antes y durante la copia.
- Antes de parsear, `OoxmlPackageInspector` inspecciona el **directorio central del ZIP sin
  descomprimir** y verifica:
  - `MaxEntryCount`: número de entradas,
  - `MaxTotalUncompressedSizeBytes`: suma de tamaños descomprimidos,
  - `MaxEntryCompressionRatio`: ratio por entrada (detección de bombas ZIP).
- Un contenedor que no es un ZIP legible produce `InvalidOfficeFileException`; superar un
  límite produce `OfficeResourceLimitExceededException`. Ambos derivan de
  `OfficeEngineException`.
- Solo se exponen límites que se **aplican de verdad**; no se anuncian límites no
  implementados (como el timeout de reloj, pendiente de un almacén con cancelación real).

### Reglas permanentes

- No se extrae contenido a rutas arbitrarias. No hay `ExtractToDirectory` ni resolución de
  rutas a partir de nombres de entrada.
- No se ejecutan macros, scripts, contenido OLE ni código incrustado.
- No se evalúan fórmulas con `eval` ni ejecución arbitraria. Las fórmulas se tratan como texto
  y, si procede, se lee su valor en caché.
- No se registra el contenido del documento en logs por defecto.
- Se distingue **protección** de hoja/libro de **cifrado** real.

### Cifrado y protección

- Un DOCX/XLSX cifrado es un contenedor OLE/CFB, no un ZIP. Hoy se rechaza como
  `InvalidOfficeFileException`. **No hay soporte de contraseñas.**
- La protección de hoja/libro (que no cifra el paquete) es metadato OOXML: se preserva al
  guardar y **no** se interpreta como seguridad; no se implementa desprotección.

## Consecuencias

- Positivas: coste de memoria acotado y rechazo temprano de entradas abusivas sin descomprimir
  datos.
- Negativas/limitaciones: los límites por defecto pueden rechazar archivos legítimos muy
  grandes; son configurables. El análisis de entradas no sustituye a un antivirus.
- Pendiente (Fase 5): límites de profundidad de XML, endurecimiento frente a entidades
  externas (verificar el comportamiento del SDK con pruebas), fuzzing y pruebas de archivos
  corruptos y maliciosos más amplias.
