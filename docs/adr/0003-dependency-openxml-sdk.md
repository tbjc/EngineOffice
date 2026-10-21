# ADR 0003 — Dependencia OOXML: Open XML SDK

- Estado: Aceptado
- Fecha: 2026-10-16
- Fase: 0

## Contexto

Se necesita leer y escribir paquetes OOXML (DOCX/XLSX) de forma fiable, y no tiene sentido
reimplementar el acceso al *Open Packaging Conventions* ni los modelos tipados de esquema.

`PROMPT.md` exige verificar la versión, la licencia y las limitaciones de cada dependencia
candidata en fuentes oficiales **antes** de decidir.

## Verificación realizada

Consultado el registro oficial de NuGet el 2026-10-16:

- Paquete: `DocumentFormat.OpenXml`.
- Última versión estable publicada: **3.5.1**.
- Licencia declarada en el `.nuspec`: `MIT` (`https://licenses.nuget.org/MIT`).
- Autor/mantenedor: Microsoft; repositorio `github.com/dotnet/Open-XML-SDK`.
- Target frameworks incluyen `net8.0`.
- Dependencia transitiva: `DocumentFormat.OpenXml.Framework` (misma versión, también MIT).

## Decisión

Adoptar `DocumentFormat.OpenXml` 3.5.1 como única dependencia de producción, fijada en
`Directory.Packages.props`.

Se acepta explícitamente su límite principal: **el SDK no compone páginas ni renderiza PDF**.
Accede al paquete y a los modelos de esquema; no es un motor de maquetación. Esto es
consistente con ADR 0005 (PDF mediante adaptador externo).

## Consecuencias

- Positivas: base sólida, mantenida y con licencia permisiva compatible con un proyecto MIT;
  validación de esquema incluida, que se usa como verificador independiente en los tests.
- Negativas: se hereda el comportamiento del SDK (por ejemplo, cómo gestiona documentos
  cifrados o relaciones malformadas) y hay que verificarlo con pruebas en vez de suponerlo.
- Obligaciones: el aviso de licencia MIT del SDK debe acompañar a la distribución
  (SBOM/aviso de terceros en la Fase 7).

## Alternativas consideradas

- Escribir el OOXML a mano sobre `System.IO.Packaging` + `XmlWriter`: descartado, mucho más
  trabajo y más riesgo de producir paquetes inválidos.
- Implementación propia de esquemas tipados: descartado por coste y mantenimiento.
