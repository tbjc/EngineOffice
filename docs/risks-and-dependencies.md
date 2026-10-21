# Riesgos, dependencias y licencias

Última actualización: 2026-10-16 (Fase 0/1).

## 1. Dependencias de producción

| Paquete | Versión | Licencia | Mantenimiento | Decisión |
| --- | --- | --- | --- | --- |
| `DocumentFormat.OpenXml` | 3.5.1 | MIT (verificado en el `.nuspec` oficial) | Microsoft; repositorio activo `dotnet/Open-XML-SDK` | Aceptada (ADR 0003) |
| `DocumentFormat.OpenXml.Framework` | 3.5.1 | MIT | transitiva del anterior | Aceptada (implícita) |

Notas sobre el SDK:

- Maneja el contenedor OPC y los modelos de esquema tipados, **no** la composición de páginas.
- Su comportamiento ante entradas malformadas o cifradas se verifica con pruebas; no se asume.
- Un documento DOCX/XLSX cifrado con contraseña es un contenedor OLE/CFB y no lo abre.

## 2. Dependencias de pruebas (no se distribuyen)

| Paquete | Versión | Licencia | Uso |
| --- | --- | --- | --- |
| `Microsoft.NET.Test.Sdk` | 17.11.1 | MIT | ejecución de pruebas |
| `xunit` | 2.9.2 | Apache-2.0 | framework de pruebas |
| `xunit.runner.visualstudio` | 2.8.2 | Apache-2.0 | adaptador de ejecución |
| `coverlet.collector` | 6.0.2 | MIT | cobertura (`--collect:"XPlat Code Coverage"`; informes en `artifacts/testresults/`) |

Todas son permisivas y compatibles con distribuir la librería bajo MIT.

## 3. Dependencias candidatas

| Candidato | Fase | Estado de la evaluación |
| --- | --- | --- |
| LibreOffice (headless) para PDF | 4 | **Adoptado como motor externo, fuera del núcleo** (ADR 0005 y ADR 0038). Verificado el 2026-10-21: ejecutable `C:\Program Files\LibreOffice\program\soffice.exe`, versión 26.8.1.1, licencia MPL-2.0 y comportamiento headless comprobado (`.docx` y `.xlsx` a PDF con perfil aislado). No es una dependencia de los paquetes de OfficeEngine: el usuario instala el motor y el adaptador lo localiza. |
| Motor externo para DOC/XLS | 6 | Sin evaluar. Sujeto a análisis de licencia y pérdidas documentadas (ADR 0006). El mismo motor sirve de punto de partida, pero la conversión heredada sigue sin implementarse. |

No se añadirá ninguna dependencia sin documentar utilidad, licencia, mantenimiento y
consecuencias de distribución (regla 9 de `AGENT.md`).

## 4. Licencia del proyecto

**Propuesta: MIT.** Motivos:

- Es compatible con todas las dependencias actuales (MIT y Apache-2.0).
- Es la licencia del SDK OOXML, lo que reduce fricción de cumplimiento.
- Permite uso comercial, que es coherente con el objetivo de ser una alternativa a librerías
  comerciales.

El archivo `LICENSE` del repositorio ya contiene el texto MIT. **Pendiente**: confirmación del
mantenedor del proyecto; si se elige otra licencia, debe revalidarse la compatibilidad.

## 5. Riesgos

| Riesgo | Impacto | Probabilidad | Mitigación | Estado |
| --- | --- | --- | --- | --- |
| Pérdida silenciosa de contenido no modelado | Alto | Media | Mutar el paquete original en vez de reconstruirlo; diagnósticos; pruebas de round-trip con contenido desconocido | Mitigado en la parte probada |
| Bomba ZIP / entrada maliciosa | Alto | Media | Inspección sin descomprimir + límites configurables | Mitigado |
| Consumo de memoria proporcional al archivo | Medio | Alta | `MaxFileSizeBytes`; almacén en streaming planificado (ADR 0004) | Aceptado temporalmente |
| Reconstrucción inconsistente de partes al guardar | Alto | Baja | Se guarda el paquete completo sin reinterpretar partes | Mitigado; vigilar en Fase 5 |
| Deriva de versiones de dependencias | Medio | Media | Central Package Management; actualización deliberada | Mitigado |
| Prometer de más (paridad con Spire/Excel) | Alto | Media | Matriz de compatibilidad con evidencia; nada sube de estado sin pruebas | Mitigado por proceso |
| Diferencias de renderizado PDF | Medio | Alta | Motor externo tras interfaz; documentar que no hay equivalencia píxel a píxel | Mitigado en la Fase 4 (ADR 0038): el motor está verificado y la diferencia está documentada |
| Incompatibilidad de DOC/XLS | Medio | Alta | Conversión externa opcional; pérdidas documentadas | Pendiente (Fase 6) |
| Comportamiento del SDK ante XML con entidades externas | Medio | Baja | Añadir pruebas de endurecimiento en la Fase 5 | Pendiente |

## 6. Decisiones pendientes

1. Confirmar la licencia final del proyecto (propuesta: MIT).
2. ~~Modelo de almacenamiento para documentos grandes y API `async`~~: **hecho** (ADR 0040: se mantiene el modelo en memoria, la API sigue siendo síncrona y la regla para adoptar `async` es el cambio del almacén).
3. Selección y validación del motor PDF: **hecho** (ADR 0005, verificado el 2026-10-21; ADR 0038).
4. ~~Alcance detallado del subconjunto de fórmulas y su matriz de funciones~~: **hecho** (ADR 0043 y `docs/formula-matrix.md`).
5. Estrategia de versionado y publicación NuGet (Fase 7).
