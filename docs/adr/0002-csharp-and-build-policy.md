# ADR 0002 — C# moderno, análisis estático y política de build

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 0

## Contexto

`AGENT.md` exige C# moderno, *nullable reference types*, análisis estático y documentación XML.
`PROMPT.md` pide que la calidad no dependa de la disciplina manual.

## Decisión

- `TargetFramework`: `net8.0` (LTS). El SDK disponible en el entorno de desarrollo es 8.0.426.
- `Nullable`: `enable` en toda la solución.
- `ImplicitUsings`: `enable`.
- `TreatWarningsAsErrors`: `true` en todos los proyectos.
- `EnableNETAnalyzers` + `AnalysisLevel=latest-recommended` + `AnalysisMode=Recommended`.
- `EnforceCodeStyleInBuild`: `true`, con `.editorconfig` versionado (incluye
  `csharp_style_namespace_declarations = file_scoped:warning`).
- `GenerateDocumentationFile`: `true` en `src/` (los avisos CS1591 son errores, luego **todo
  miembro público está documentado**). Desactivado en tests.
- `Deterministic`: `true`, `ContinuousIntegrationBuild` cuando `CI=true`.
- Central Package Management (`Directory.Packages.props`) con versiones fijadas.
- Los nombres con guion bajo (`CA1707`) se permiten solo en el proyecto de tests.

## Consecuencias

- Un cambio que introduce un aviso rompe el build: la deuda deja de acumularse en silencio.
- Documentar cada miembro público es obligatorio; encarece un poco añadir API, a cambio de
  mantenerla intencionada.
- Fijar versiones evita que una actualización de dependencia cambie el build sin revisión.
  Actualizar una dependencia es una acción deliberada (ver ADR 0003).
