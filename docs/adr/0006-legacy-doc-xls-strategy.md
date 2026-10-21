# ADR 0006 — Formatos heredados DOC/XLS

- Estado: Aceptado (diseño); implementación pendiente (Fase 6)
- Fecha: 2026-10-16
- Fase: 0 → 6

## Contexto

DOC y XLS son formatos binarios compuestos (OLE/CFB) del ecosistema previo a OOXML. No son
DOCX/XLSX renombrados: tienen estructura, semántica y fidelidad distintas.

## Decisión

1. **No** se promete lectura/escritura nativa de DOC/XLS en las fases iniciales.
2. Se distinguen y se documentan como capacidades **no equivalentes**:
   - *convertir* a DOCX/XLSX,
   - *leer* el formato original,
   - *editar* el formato original,
   - *guardar en el formato original*.
   Un adaptador puede ofrecer la primera y no las demás.
3. La vía preferida es la conversión con un motor externo (por ejemplo LibreOffice headless) a
   DOCX/XLSX, y después usar el núcleo moderno.
4. Se añadirá soporte solo con: fixtures reales de varias versiones y complejidad, pruebas de
   regresión, análisis de licencia del motor y documentación explícita de las pérdidas.

## Consecuencias

- Positivas: no se genera una promesa falsa de compatibilidad ni se distrae esfuerzo del núcleo
  DOCX/XLSX.
- Negativas: algunos usuarios no encontrarán soporte DOC/XLS durante mucho tiempo; es una
  limitación de producto aceptada conscientemente.
- Nota: abrir un DOC/XLS con las API actuales produce `InvalidOfficeFileException` porque el
  contenedor no es un paquete OOXML (ZIP). Es el comportamiento correcto, no un fallo.
