# ADR 0020 — Propiedades básicas del documento

- Estado: Aceptado
- Fecha: 2026-10-19
- Fase: 2, incremento 2.10

## Contexto

Un `.docx` guarda metadatos en varias partes. Las **propiedades básicas** que muestra Word en la
ficha del documento viven en `docProps/core.xml` como elementos de los namespaces `dc`, `cp`,
`dcterms` y `dcmitype`: título, asunto, autor, palabras clave, descripción, último editor, revisión,
fechas de creación/modificación/impresión, categoría, estado, identificador, idioma y versión.

El paquete también puede contener **propiedades extendidas** (`docProps/app.xml`: empresa, gerente,
aplicación, estadísticas…) y **propiedades personalizadas** (`docProps/custom.xml`). Hasta este
incremento todas se preservaban porque el SDK conserva el paquete, pero la API no las leía ni las
editaba.

El criterio del incremento es que las propiedades básicas sobrevivan al round-trip y que las
propiedades no modeladas no se pierdan.

## Decisión

### 1. Alcance: propiedades básicas de `docProps/core.xml`

- `WordDocument.Properties` devuelve un `DocumentProperties`, una **vista viva** del documento.
- `DocumentProperties` expone las quince propiedades básicas: `Title`, `Subject`, `Creator`,
  `Keywords`, `Description`, `LastModifiedBy`, `Revision`, `LastPrinted`, `Created`, `Modified`,
  `Category`, `ContentStatus`, `Identifier`, `Language` y `Version`.
- Leer una propiedad no almacenada devuelve `null`; asignar `null` la elimina.
- Las propiedades extendidas y personalizadas **no** se modelan. No se tocan, así que se conservan
  al guardar.

`ContentType` no se expone: describe la parte de propiedades, no un metadato editable del documento.
`Revision` se expone como texto porque OOXML lo define como `xsd:string`; no se convierte a número ni
se valida como entero.

### 2. La parte se crea en `docProps/core.xml`

El SDK ofrece `WordprocessingDocument.PackageProperties`, pero en 3.5.1 crea la parte de propiedades
OPC en una ruta arbitraria (`package/services/metadata/core-properties/<guid>.psmdcp`) cuando el
documento no tiene una. Para que el paquete producido use la convención de Word, al asignar el primer
valor se crea un `CoreFilePropertiesPart` y se inicializa su raíz vacía en `docProps/core.xml`.

Consecuencias:

- Un documento nuevo sin propiedades **no** crea `docProps/core.xml`; leer las propiedades tampoco.
- Asignar `null` a una propiedad inexistente tampoco crea la parte.
- Si el documento ya tiene una parte de core properties (en la ruta convencional o en otra), se lee
  y se edita **en su sitio**; no se duplica ni se mueve.

### 3. El SDK sigue siendo el lector y escritor

El SDK 3.5.1 expone las propiedades básicas mediante `IPackageProperties`, una interfaz marcada como
experimental (`OOXML0001`). Se usa de forma encapsulada dentro de `DocumentProperties`, sin exponerla
en la API pública y sin nombrar el tipo en el código de producción. La versión del SDK está fijada y
las pruebas cubren el round-trip completo. Si una versión futura cambia esa interfaz, solo debe
cambiar esta clase.

No se reescribe `docProps/core.xml` a mano: el SDK conoce los nombres de elemento, los tipos
`dcterms:W3CDTF` y las reglas de serialización. Eso reduce el riesgo de generar marcado inválido.

### 4. Fechas

`Created`, `Modified` y `LastPrinted` se almacenan como un instante con offset. La API usa
`DateTime?` porque es el tipo que expone el SDK. Al leer, el SDK devuelve un `DateTime` local; el
instante se conserva, pero el offset textual puede cambiar al volver a guardar. Para comparar valores
hay que usar `ToUniversalTime()`.

`Created` y `Modified` se serializan con `xsi:type="dcterms:W3CDTF"`. `LastPrinted` es un
`xsd:dateTime` de `cp` y no lleva ese atributo.

### 5. Validación de texto

Los valores de texto pueden contener tabuladores y saltos de línea, pero deben ser caracteres XML
válidos. Antes de asignar se valida con `XmlConvert.VerifyXmlChars`; un carácter imposible de
representar (por ejemplo `U+0000` o un surrogate suelto) lanza `ArgumentException` y no se crea ni se
modifica la parte.

### 6. Lectura estricta

Si `docProps/core.xml` existe pero no se puede interpretar (por ejemplo una fecha mal formada), la
lectura lanza `InvalidOfficeFileException` en vez de dejar escapar una `XmlException` o inventar un
valor. El documento no se reescribe en esa lectura.

## Consecuencias

- API pública nueva: `WordDocument.Properties` y `DocumentProperties`.
- Se modelan las propiedades básicas; las extendidas y personalizadas se preservan pero no se leen
  ni se editan.
- La parte `docProps/core.xml` se crea solo cuando hace falta y en la ruta convencional.
- El offset textual de las fechas puede cambiar; el instante no.
- La dependencia de una interfaz experimental del SDK queda aislada en una sola clase.
- No se modelan propiedades personalizadas, estadísticas de `app.xml`, empresa/gerente ni
  `ContentType`.

## Verificación

- `WordPropertiesTests`: documento nuevo sin propiedades ni parte; valores completos por archivo;
  edición por stream; borrado; preservación de `app.xml`/`custom.xml`; tabuladores y saltos;
  rechazo de caracteres XML inválidos; fecha mal formada con error tipado; uso tras `Dispose`.
- `WordPropertiesInteropTests`: paquete válido con `OpenXmlValidator`; lectura con el SDK;
  nombres/namespaces esperados en el XML; edición de un fixture ajeno sin perder partes no
  modeladas; creación de `docProps/core.xml` al editar un documento que no lo tenía.
- Fixtures construidos con el SDK (core, extended y custom) leídos y editados por OfficeEngine.
