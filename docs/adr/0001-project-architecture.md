# ADR 0001 — Arquitectura por capas y superficie pública

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 0

## Contexto

OfficeEngine debe cubrir Word y Excel sobre OOXML, exponer una API propia y mantenible, y no
acoplar a los usuarios a las clases del SDK oficial. Además la conversión a PDF y el soporte
de formatos heredados implican motores externos que no deben mezclarse con el núcleo.

## Decisión

Dividir la solución en capas con dependencias unidireccionales:

```
OfficeEngine.Core      → contratos, opciones, límites, diagnósticos, errores tipados
OfficeEngine.OpenXml   → infraestructura OOXML interna (internal + InternalsVisibleTo)
OfficeEngine.Word      → WordprocessingML (.docx)
OfficeEngine.Excel     → SpreadsheetML (.xlsx)
OfficeEngine.Tests     → pruebas
```

Reglas:

1. `Core` no referencia ningún formato ni el SDK OOXML.
2. `OpenXml` concentra lo común a Word y Excel (lectura acotada, inspección de paquete) y no
   conoce ningún modelo de dominio.
3. `Word` y `Excel` no exponen tipos del SDK como parte obligatoria de su API pública.
4. PDF, fórmulas y formatos heredados son subsistemas separados, con proyecto propio y
   añadidos solo cuando su fase comienza.

## Consecuencias

- Positivas: la API pública es pequeña y estable; el SDK puede actualizarse o sustituirse sin
  romper contratos; las pruebas pueden verificar cada capa por separado.
- Negativas: algo de indirección y código de "pegamento"; hay que evitar que `OpenXml`
  acumule lógica de dominio.
- Neutras: no se crean proyectos vacíos para capas futuras; su estructura está documentada en
  `docs/architecture/overview.md`.
