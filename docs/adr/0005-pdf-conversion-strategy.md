# ADR 0005 - Estrategia de conversión a PDF

- Estado: Aceptado (diseño); implementado en la Fase 4 (ADR 0038)
- Fecha: 2026-10-16
- Fase: 0 → 4

## Contexto

La exportación a PDF es un requisito del producto, pero OOXML no describe la composición de
páginas. El SDK oficial no renderiza; se necesita un motor de maquetación.

## Decisión

1. Definir en la Fase 4 una interfaz `IPdfConverter` en un proyecto `OfficeEngine.Pdf`,
   separada del núcleo, con opciones, resultado, límites y diagnósticos propios.
2. Implementar un primer adaptador **LibreOffice headless** en `OfficeEngine.Pdf.LibreOffice`,
   **solo tras** verificar en el entorno: disponibilidad del ejecutable, versión mínima
   soportada, requisitos del sistema, comportamiento headless, dependencias de fuentes y
   licencia de distribución.
3. El adaptador debe verificar la presencia y versión del ejecutable antes de convertir,
   usar directorios temporales aislados por ejecución, aplicar timeout, límite de tamaño y
   concurrencia, capturar la salida de diagnóstico **sin registrar el contenido del
   documento**, y detectar códigos de salida fallidos, archivos ausentes y PDF inválidos.
4. La exportación debe partir del **estado en memoria** del documento, nunca de una copia
   antigua en disco. La estrategia concreta (volcado temporal del estado actual y marca de
   versión) se define con pruebas en la Fase 4.
5. Las pruebas que requieran el motor externo se marcan como **pruebas de integración** y se
   **omiten con un motivo explícito** cuando el motor no está instalado.

## Estado del entorno

### Comprobado el 2026-10-16

LibreOffice **no estaba instalado** en el entorno de desarrollo
(`C:\Program Files\LibreOffice\program\soffice.exe` no existía). Por tanto:

- No se había validado ningún adaptador.
- Las futuras pruebas de integración de PDF se omitirían en este entorno y lo declararían.
- **No** se afirma ningún soporte de PDF en la Fase 1.

### Comprobado el 2026-10-21

El motor **sí está instalado** y se ha verificado lo que este ADR pedía:

| Comprobación | Resultado |
| --- | --- |
| Ejecutable | `C:\Program Files\LibreOffice\program\soffice.exe` (y `soffice.com`, la variante de consola) |
| Versión | 26.8.1.1 |
| Licencia | MPL-2.0 (`license.txt`) |
| Headless | Convierte `.docx` y `.xlsx` en PDF con `--convert-to`, con perfil aislado |

La implementación de esta estrategia (contrato, aislamiento, límites, comprobaciones y manejo de
errores) está en `docs/adr/0038-pdf-libreoffice-converter.md`, y el resultado real, en
`docs/phase-4-report.md`.

## Consecuencias

- Positivas: el núcleo permanece libre de dependencias de motores; el motor es sustituible; el
  proceso queda aislado y observable.
- Negativas/limitaciones a documentar: la fidelidad no puede garantizarse igual a Microsoft
  Office; depende de fuentes, versión, configuración de página y sistema operativo. Esto debe
  comunicarse explícitamente y no se promete equivalencia píxel a píxel.
