# Matriz de compatibilidad de OfficeEngine

Leyenda:

- **No iniciado**: no existe implementación.
- **En progreso**: parcial, sin cubrir el alcance completo.
- **Implementado**: existe implementación, sin prueba dedicada que la respalde.
- **Probado**: implementación con pruebas automatizadas que pasan.
- **Limitación conocida**: comportamiento documentado que no debe interpretarse como soporte
  completo.

> Regla: una capacidad no sube a *Implementado* o *Probado* sin evidencia real. El contenedor
> ZIP poder abrirse **no** cuenta como prueba.

Última actualización: 2026-10-21 (tras los incrementos 6.1 y 6.2 de la Fase 6).

## Núcleo (OfficeEngine.Core)

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Errores tipados (`OfficeEngineException` y derivadas) | Probado | `ResourceLimitTests.TypedExceptions_DeriveFromOfficeEngineException` |
| Límites de recursos configurables y validación | Probado | `OfficeEngineLimitsTests` |
| Diagnósticos y severidad | Probado | `OfficeEngineDiagnosticsTests` |
| Color RGB (`RgbColor`: constructor, hex, igualdad) | Probado | `RgbColorTests` |

## Infraestructura OOXML (OfficeEngine.OpenXml, interna)

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Carga acotada de archivo/stream | Probado | `ResourceLimitTests.*_Open_ExceedingMaxFileSize_*` |
| Inspección del contenedor sin descomprimir | Probado | `ResourceLimitTests.*_CompressionBomb_*`, `*_ExceedingMaxEntryCount_*` |
| Rechazo de entradas que no son paquetes OOXML | Probado | `ResourceLimitTests.*_ZipThatIsNot*` |

## Word — DOCX (OfficeEngine.Word)

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Crear documento | Probado | `WordDocumentTests.Create_StartsEmpty` |
| Abrir desde archivo | Probado | `WordDocumentTests.Save_ToFile_ThenOpen_RoundTrips` |
| Abrir desde stream (+ `leaveOpen`) | Probado | `WordDocumentTests.Open_FromStream_*` |
| Guardar a archivo (atómico) | Probado | `WordDocumentTests.Save_ToFile_ThenOpen_RoundTrips` |
| Guardar a stream | Probado | `WordDocumentTests.Create_SaveToStream_Reopen_RoundTripsParagraphText` |
| Guardados sucesivos conservan las ediciones | Probado | `WordDocumentTests.Save_CanBeCalledRepeatedly_AndKeepsEveryEdit` |
| Párrafos: lectura y escritura de texto | Probado | `WordDocumentTests.AddParagraph_ReturnsEditableParagraph` |
| Preservación de espacios significativos | Probado | `WordDocumentTests.Create_PreservesWhitespaceInParagraphText` |
| Preservar propiedades de párrafo al reemplazar el texto | Probado | `WordCharacterFormatTests.Paragraph_TextSetter_PreservesParagraphPropertiesAndDropsTheOldRuns` |
| Conservación de contenido no modelado | Probado | `WordInteropTests.OpenedDocument_WithUnmodelledBodyContent_ReportsDiagnosticAndPreservesIt` |
| Diagnóstico de contenido no modelado | Probado | idem; una tabla modelada ya no genera diagnóstico (`WordInteropTests.OpenedDocument_WithModelledContent_HasNoDiagnostics`) |
| Validación del paquete producido (SDK externo) | Probado | `WordInteropTests.ProducedPackage_IsAValidWordprocessingDocument` y `WordFormattingInteropTests.ProducedDocument_WithCharacterFormatting_IsSchemaValid` |
| Runs: inserción y lectura en orden | Probado | `WordCharacterFormatTests.AddRun_AppendsRunsAndPreservesTheirOrder` |
| Formato de carácter: familia y tamaño de fuente | Probado | `WordCharacterFormatTests.Run_CharacterFormatting_RoundTripsThroughSaveAndOpen`, `WordFormattingInteropTests.ProducedDocument_CharacterFormatting_IsReadableByTheOpenXmlSdk` |
| Formato de carácter: negrita y cursiva (estado heredado/explicitado) | Probado | `WordCharacterFormatTests.Run_TriStateFormatting_DistinguishesUnsetFromExplicitlyDisabled` |
| Formato de carácter: subrayado (los 18 valores `ST_Underline`) | Probado | `WordCharacterFormatTests.Run_AllUnderlineStyles_RoundTrip` |
| Formato de carácter: color RGB explícito | Probado | `WordFormattingInteropTests.ProducedDocument_CharacterFormatting_IsReadableByTheOpenXmlSdk` |
| Orden de esquema de `w:rPr` en el paquete producido | Probado | `WordFormattingInteropTests.ProducedDocument_WritesRunPropertiesInSchemaOrder` |
| Preservar propiedades de run no modeladas al editar otras | Probado | `WordFormattingInteropTests.OpenedDocument_UnmodelledRunProperties_SurviveEditingModelledOnes` |
| Formato de carácter: resaltado, tachado y resto de `w:rPr` | Limitación conocida | No se modelan ni se editan; **se preservan** y hay prueba de ello |
| Color de tema (`w:themeColor`) o `auto` | Limitación conocida | No tiene RGB fijo, se devuelve como `null` y el marcado original se conserva (`WordFormattingInteropTests.OpenedDocument_WithThemeColor_*`) |
| Fuentes/tamaños de escritura compleja (`w:rFonts` de Asia y cs, `w:szCs`) | Limitación conocida | No se modelan; se conservan sin cambios |
| Archivo ilegible: tamaño de fuente no numérico | Probado | `WordFormattingInteropTests.OpenedDocument_WithMalformedFontSize_ThrowsInvalidOfficeFileException` |
| Archivo ilegible: valor on/off inválido | Probado | `WordFormattingInteropTests.OpenedDocument_WithMalformedOnOffValue_ThrowsInvalidOfficeFileException` |
| Archivo ilegible: estilo de subrayado desconocido | Probado | `WordFormattingInteropTests.OpenedDocument_WithUnknownUnderlineStyle_ThrowsInvalidOfficeFileException` |
| Alineación de párrafo (`w:jc`): `Left`, `Center`, `Right`, `Justify` (`distribute`) y `Both` | Probado | `WordParagraphFormatTests.Paragraph_Alignment_AllValuesRoundTrip`, `WordParagraphFormatInteropTests.ProducedDocument_ParagraphFormatting_IsReadableByTheOpenXmlSdk` |
| Sangría de párrafo (`w:ind`): primera línea, izquierda, derecha y francesa, en puntos | Probado | `WordParagraphFormatTests.Paragraph_Indents_RoundTrip`, `...Paragraph_Indents_KeepUnmodelledAttributes` |
| Espaciado de párrafo (`w:spacing`): anterior, posterior y línea exacta, en puntos | Probado | `WordParagraphFormatTests.Paragraph_Spacing_RoundTrip`, `...Paragraph_StoredMultipleLineSpacing_ThrowsAndIsPreserved` |
| Orden de esquema de `w:pPr` en el paquete producido | Probado | `WordParagraphFormatInteropTests.ProducedDocument_WritesParagraphPropertiesInSchemaOrder`, `...OpenedDocument_WithNumbering_KeepsItAheadOfTheModelledProperties`, `...OpenedDocument_WithUnmodelledParagraphFlags_KeepsSchemaOrder` |
| Preservar propiedades de párrafo no modeladas al editar otras | Probado | `WordParagraphFormatInteropTests.OpenedDocument_WithMixedParagraphFormatting_PreservesUnmodelledMarkup`, `WordParagraphFormatTests.Paragraph_Indents_KeepUnmodelledAttributes` |
| Alineación no modelada (`start`, `end`, kashida...) | Limitación conocida | Es un valor válido de `ST_Jc` que no se modela: se lanza `UnsupportedOfficeFeatureException` y el marcado se conserva (`Paragraph_UnmodelledStoredAlignment_ThrowsAndIsPreserved`) |
| Espaciado de línea como múltiplo (`w:lineRule="auto"`) | Limitación conocida | No es expresable en puntos: `LineSpacingPoints` lanza `UnsupportedOfficeFeatureException` en lectura; al asignar se escribe `w:lineRule="exact"` (`Paragraph_StoredMultipleLineSpacing_ThrowsAndIsPreserved`) |
| Sangría en caracteres (`w:leftChars`, `w:firstLineChars`) | Limitación conocida | No se modela; se preserva y no se reporta por las propiedades en puntos (`Paragraph_Indents_KeepUnmodelledAttributes`) |
| Archivo ilegible: alineación, sangría, espaciado o regla de línea mal formados | Probado | `WordParagraphFormatTests.Paragraph_MalformedStored*_ThrowsInvalidOfficeFile` |
| Rango de los valores en puntos (no desbordar los twips) | Probado | Un valor que no cabe en un twips de 32 bits se rechaza con `ArgumentOutOfRangeException` en vez de desbordar (`Paragraph_RejectsInvalidNumericValues`) |
| Tabulaciones de párrafo (`w:tabs`): posición, alineación (`left`, `center`, `right`, `decimal`, `bar`, `clear`) y relleno (`dot`, `hyphen`, `underscore`, `heavy`, `middleDot`) | Probado | `WordTabStopTests.Paragraph_AddTabStop_WritesEveryModelledAlignmentAndLeader`, `WordTabAndBreakInteropTests.ProducedDocument_TabStops_AreReadableByTheOpenXmlSdk` |
| Tabulación y saltos explícitos en un run (`w:tab`, `w:br` de texto/página/columna) | Probado | `WordRunContentTests.Run_AddTab_AppendsATabCharacter`, `...Run_AddBreak_AppendsEachModelledBreakType`, `WordTabAndBreakInteropTests.ProducedDocument_RunContent_IsReadableByTheOpenXmlSdk` |
| Contenido modelado del run en orden de documento (`Run.Content`) | Probado | `WordRunContentTests.Run_Content_ReportsContentInDocumentOrder`, `...Run_Content_RoundTripsThroughSaveAndOpen` |
| Retorno de carro ajeno (`w:cr`) leído como salto de línea | Probado | `WordRunContentTests.Run_Content_MapsACarriageReturnToATextWrappingBreak` |
| Rechazo de tabuladores y saltos en las API de texto plano (se construyen con `Run.AddTab`/`Run.AddBreak`) | Probado | `WordRunContentTests.Run_AddText_RejectsTabsAndLineBreaks`, `...Paragraph_TextSetter_RejectsTabsAndLineBreaks`, `...WordDocument_AddParagraph_RejectsTabsAndLineBreaks` |
| Rechazo de caracteres que Office Open XML no puede almacenar (p. ej. `U+0000` o un suplente suelto) | Probado | `WordRunContentTests.Run_AddText_RejectsCharactersOfficeOpenXmlCannotStore`, `...Run_AddText_RejectsALoneSurrogate` |
| Alineación de tabulación no modelada (`start`, `end`, `num`) | Limitación conocida | Son valores válidos de `ST_TabStop` que no se modelan: se lanza `UnsupportedOfficeFeatureException` y el marcado se conserva (`Paragraph_StoredUnmodelledTabStop_ThrowsAndIsPreserved`) |
| Archivo ilegible: tabulación sin `w:val`/`w:pos`, con posición o alineación mal formadas, o tipo de salto desconocido | Probado | `WordTabStopTests.Paragraph_TabStops_Missing*_ThrowsInvalidOfficeFile`, `...Paragraph_TabStops_Malformed*_ThrowsInvalidOfficeFile`, `WordRunContentTests.Run_Content_MalformedStoredBreakType_ThrowsInvalidOfficeFile` |
| `w:clear` de `w:br` | Limitación conocida | No se modela; se preserva al guardar pero no se reporta (`RunBreak`) |
| Estilos de párrafo y carácter | En progreso | Referencias de estilo y catálogo de solo lectura implementados y probados (filas siguientes); no se crean ni editan definiciones |
| Referencia al estilo de párrafo (`Paragraph.StyleId`, `w:pStyle`) | Probado | `WordStyleReferenceTests.Paragraph_StyleId_*`, `WordStyleInteropTests.ProducedDocument_StyleReferences_AreReadableByTheOpenXmlSdk` |
| Referencia al estilo de carácter (`Run.StyleId`, `w:rStyle`) | Probado | `WordStyleReferenceTests.Run_StyleId_*` |
| Catálogo de estilos de solo lectura (`WordDocument.Styles`, `FindStyle`) | Probado | `WordStyleCatalogTests.Styles_*`, `...FindStyle_*` |
| Formato que aporta un estilo (carácter y párrafo) | Probado | `WordStyleCatalogTests.OwnFormatting_*`, `WordStyleInteropTests.ForeignDocument_WithStyles_IsReadAfterEdit` |
| Herencia de estilos resuelta (`w:basedOn`) | Probado | `WordStyleCatalogTests.EffectiveFormatting_*` |
| Ciclo en `w:basedOn` (incluida la autorreferencia) | Probado | No se puede resolver: `InvalidOfficeFileException` (`EffectiveFormatting_Throws_WhenTheChainHasACycle`, `...WhenAStyleIsBasedOnItself`) |
| Crear/editar/eliminar definiciones de estilo | No iniciado | El catálogo es de solo lectura y `styles.xml` se conserva byte a byte (`WordStyleInteropTests.StylesPart_IsPreservedWhenTheBodyIsEdited`) |
| `w:docDefaults`, estilo por defecto, `w:link` y `w:next` | Limitación conocida | `EffectiveFormatting` resuelve solo la cadena `w:basedOn`; no los pliega (ADR 0011) |
| Propiedades de estilo no modeladas e interlineado de estilo | Limitación conocida | No se reportan por `OwnFormatting`/`EffectiveFormatting`; se preservan. El modelo cubre fuente, tamaño, negrita, cursiva, color, alineación, sangrías y espaciado anterior/posterior |
| Referencia de estilo colgante (`w:pStyle`/`w:rStyle` a un estilo inexistente) | Limitación conocida | Válida para el esquema y para Word: se guarda y se lee tal cual, no se reporta (`WordStyleInteropTests.DanglingStyleReference_RoundTrips`) |
| Archivo ilegible: `w:styleId`/`w:name`/`w:basedOn`/`w:type`/`w:default` con valor ausente o mal formado | Probado | `WordStyleCatalogTests.Style_*_Throws*`, `...OwnFormatting_Throws*` |
| Listas numeradas y con viñetas | En progreso | Referencias de lista, catálogo de solo lectura y creación de listas de un nivel implementados y probados (filas siguientes); no se personalizan niveles |
| Referencia a la lista y al nivel de un párrafo (`Paragraph.ListId`/`ListLevel`, `w:numPr`) | Probado | `WordListReferenceTests.Paragraph_ListId_*`, `...Paragraph_ListLevel_*`, `...Paragraph_ListId_AndLevel_AreInSchemaOrderWithinTheNumberingProperties` |
| `Paragraph.ApplyList` para situar un párrafo en una lista y un nivel | Probado | `WordListReferenceTests.Paragraph_ApplyList_*` |
| Catálogo de listas de solo lectura (`WordDocument.Lists`, `FindList`) | Probado | `WordListCatalogTests.Lists_*`, `...FindList_*` |
| Niveles de una lista (`w:lvl`): formato, texto, inicio y justificación | Probado | `WordListCatalogTests.List_Level_*`, `WordListInteropTests.ForeignDocument_ListLevelsAreReadFromTheRawMarkup` |
| Crear una lista de un nivel (`WordDocument.AddList`), numerada o con viñetas | Probado | `WordListCatalogTests.AddList_*`, `WordListInteropTests.ProducedDocument_WithLists_IsSchemaValid` |
| Orden de esquema de `w:numPr`, de los hijos del nivel y de la parte de numeración | Probado | `WordListReferenceTests.Paragraph_ListId_IsInsertedInSchemaOrder`, `WordListCatalogTests.AddList_AppendsANumberedLevelInSchemaOrder`, `...AddList_WritesEveryAbstractNumberingBeforeTheInstances` |
| Editar el formato del párrafo o el cuerpo no rompe la lista | Probado | `WordListInteropTests.ParagraphFormattingEdits_DoNotBreakTheList`, `...NumberingPart_IsPreservedWhenTheBodyIsEdited` |
| `numbering.xml` se conserva al editar el cuerpo y al añadir una lista | Probado | `WordListInteropTests.NumberingPart_IsPreservedWhenTheBodyIsEdited`, `...ExistingNumberingIsPreserved_WhenANewListIsAdded` |
| Formato de numeración no modelado (p. ej. `chicago`) | Limitación conocida | Es un valor válido de `ST_NumberFormat` que no se modela: se lanza `UnsupportedOfficeFeatureException` y el marcado se conserva (`WordListCatalogTests.List_Level_Format_Throws_WhenTheValueIsNotModelled`) |
| Personalizar niveles (texto, inicio, reinicios, sufijo), `w:numPicBullet`, numeración ligada a estilos y editar/eliminar definiciones de lista | Limitación conocida | La API crea listas de un solo nivel y lee el catálogo; no edita definiciones (ADR 0012) |
| Propiedades de párrafo/carácter de un nivel (`w:pPr`/`w:rPr` del nivel) | Limitación conocida | No se modelan; se preservan, pero no se reportan |
| Referencia de lista colgante (`w:numId` a una lista inexistente, o sin `numbering.xml`) | Limitación conocida | Válida para el esquema y para Word: se guarda y se lee tal cual; `Levels` queda vacío (`WordListCatalogTests.List_Levels_IsEmpty_WhenTheAbstractNumberingIsMissing`, `WordListInteropTests.DanglingListReference_RoundTrips`) |
| Archivo ilegible: `w:numId`/`w:ilvl`/`w:abstractNumId`/`w:numFmt`/`w:lvlJc` con valor ausente o mal formado | Probado | `WordListReferenceTests.Paragraph_*_Throws*`, `WordListCatalogTests.List_*_Throws*` |
| Tablas: estructura y contenido (crear/leer/editar filas, celdas y su texto) | Probado | `WordTableTests`, `WordTableInteropTests.ProducedDocument_WithTable_IsSchemaValid` |
| Tabla como elemento de bloque, intercalada con párrafos | Probado | `WordTableTests.AddTable_AppendsTheTableAtTheEndOfTheBody`, `WordTableInteropTests.ProducedDocument_InterleavesParagraphsAndTables` |
| Crear una tabla (`WordDocument.AddTable(filas, columnas)`) en orden de esquema | Probado | `WordTableTests.AddTable_*`, `WordTableInteropTests.ProducedDocument_WritesTheTableChildrenInSchemaOrder`, `...DeclaresOneGridColumnPerColumn` |
| Crecer una tabla con una fila (`Table.AddRow`) | Probado | `WordTableTests.Table_AddRow_*` |
| Texto de una celda (`TableCell.Text`, `AddParagraph`) y del documento con celdas | Probado | `WordTableTests.TableCell_*`, `WordTableTests.Text_*`, `WordTableInteropTests.ProducedDocument_TableText_IsReadableByTheOpenXmlSdk` |
| Preservar el marcado de tabla no modelado al editar el contenido | Probado | `WordTableInteropTests.OpenedDocument_WithFormattedTable_KeepsUnmodelledMarkupWhenTheContentIsEdited`, `...WithNestedTable_PreservesTheNestedTableWhenTheOuterContentIsEdited` |
| Columnas fijas al crear la tabla | Limitación conocida | No se añaden ni se quitan columnas; `Table.AddRow` usa la rejilla existente (ADR 0013) |
| Tabla anidada en una celda | Limitación conocida | No se modela ni la devuelve `Tables`/`TableCell.Paragraphs`, y su texto no entra en `WordDocument.Text`; se preserva al editar (`WordTableTests.Text_DoesNotIncludeTextOfANestedTable`) |
| Propiedades de tabla: ancho preferido, alineación y disposición (`Table.Width`/`Alignment`/`Layout`) | Probado | `WordTableFormattingTests.Table_WidthPoints_*`, `...Table_Width_*`, `...Table_Alignment_*`, `...Table_Layout_*`, `WordTableFormattingInteropTests.ProducedDocument_WithTableAndCellFormatting_IsSchemaValid` |
| Propiedades de celda: ancho y alineación vertical (`TableCell.Width`/`VerticalAlignment`) | Probado | `WordTableFormattingTests.TableCell_WidthAndVerticalAlignment_*`, `...TableCell_VerticalAlignment_*` |
| Bordes de tabla y de celda (`Table.Border*`, `TableCell.Border*`, `TableBorder`) | Probado | `WordTableFormattingTests.Table_Borders_*`, `...Table_Border_*`, `...TableCell_Borders_*`, `WordTableFormattingInteropTests.ProducedDocument_EveryBorderStyle_IsSchemaValidAndReadable` |
| Sombreado de tabla y de celda (`Table.ShadingColor`, `TableCell.ShadingColor`) | Probado | `WordTableFormattingTests.Table_ShadingColor_*`, `...TableCell_ClearingTheLastProperty_*` |
| Orden de esquema de `w:tblPr`/`w:tcPr`/`w:tblBorders`/`w:tcBorders` al escribir propiedades | Probado | `WordTableFormattingTests.Table_Formatting_IsInsertedInSchemaOrder*`, `...TableCell_Formatting_IsInsertedInSchemaOrder*`, `WordTableFormattingInteropTests.ProducedDocument_WritesTablePropertiesInSchemaOrder`, `...WritesCellPropertiesInSchemaOrder` |
| Editar las propiedades no rompe el marcado no modelado de la tabla (estilo, márgenes de celda, bordes internos, `w:tblLook`) | Probado | `WordTableFormattingInteropTests.OpenedDocument_WithFormattedTable_KeepsUnmodelledMarkupWhenFormattingIsEdited` |
| Valor de formato definido por OOXML y no modelado (unidad `nil`, patrón de sombreado distinto de `clear`, borde decorativo) | Limitación conocida | Falla con `UnsupportedOfficeFeatureException` y el marcado se conserva (`WordTableFormattingTests.Table_Width_ReportsAnUnmodelledUnit`, `...Table_ShadingColor_ReportsAnUnmodelledPattern`, `...Table_Border_ReportsUnmodelledOrMalformedMarkup`) |
| Archivo ilegible: `w:w`/`w:sz`/`w:color`/`w:val` de tabla mal formado, ausente o negativo | Probado | `WordTableFormattingTests.Table_Width_ReportsAMalformedMeasure`, `...Table_Border_ReportsUnmodelledOrMalformedMarkup`, `...Table_Border_ReportsAMalformedColour`, `...Table_Alignment_ReportsAnUnmodelledValue` |
| Márgenes de celda (`w:tblCellMar`/`w:tcMar`) | Limitación conocida | No se modelan; se preservan al editar (ADR 0014) |
| Combinación de celdas: span horizontal (`TableCell.ColumnSpan`, `w:gridSpan`) | Probado | `WordTableMergeTests.Cell_ColumnSpan_*`, `WordTableMergeInteropTests.ProducedDocument_WritesTheMergedRowWithTheFewerCellsOfAMergedRow` |
| Combinación de celdas: combinación vertical (`TableCell.VerticalMerge`, `w:vMerge`) | Probado | `WordTableMergeTests.Cell_VerticalMerge_*`, `WordTableMergeInteropTests.ProducedDocument_EveryVerticalMergeState_IsSchemaValidAndReadable` |
| Combinar y dividir celdas de una fila (`TableRow.MergeCells`/`UnmergeCell`), moviendo el contenido a la celda que absorbe | Probado | `WordTableMergeTests.MergeCells_*`, `...UnmergeCell_*`, `WordTableMergeInteropTests.ProducedDocument_AfterSplittingAMergedCell_HasOneCellPerGridColumn` |
| Combinar y dividir un tramo de filas (`Table.MergeCellsVertically`/`UnmergeCellsVertically`), solo con las marcas | Probado | `WordTableMergeTests.MergeCellsVertically_*`, `...UnmergeCellsVertically_*` |
| El span de una fila sigue sumando las columnas de la rejilla | Probado | `WordTableMergeInteropTests.ProducedDocument_WithMergedCells_IsSchemaValid` |
| Combinar celdas cuyo ancho no se puede sumar (falta, tipos mezclados, `auto`) deja la celda combinada sin ancho preferido | Probado | `WordTableMergeTests.MergeCells_WithWidthsThatCannotBeAddedUp_StatesNoWidth`, `...MergeCells_AddsUpWidthsOfTheSameKind` |
| Combinaciones rectangulares (una combinación horizontal que forma parte de una vertical) | Limitación conocida | Se rechazan con `InvalidOperationException` sin escribir nada (`WordTableMergeTests.MergeCells_OfACellOfAVerticalMerge_Throws`, `...UnmergeCell_OfACellOfAVerticalMerge_Throws`, `...MergeCellsVertically_WhenTheCellsCoverDifferentSpans_Throws`) |
| `w:hMerge` heredado y span mal formado (sin `w:val`, no entero, `< 1`) | Limitación conocida | `w:hMerge` no se modela y se preserva; un span mal formado falla con `InvalidOfficeFileException` (`WordTableMergeTests.Cell_ColumnSpan_StoredValueThatIsNotAColumnCount_IsRejected`, `WordTableMergeInteropTests.OpenedDocument_WithMergeMarkupThisVersionDoesNotModel_KeepsItWhenMerging`) |
| Eliminar o mover tablas, filas o celdas | Limitación conocida | Solo se añaden al final y la combinación absorbe las celdas que cubre; borrar/mover a mano sigue fuera de alcance, igual que en párrafos |
| Secciones: lectura en orden de documento (`WordDocument.Sections`) | Probado | `WordSectionTests.Sections_ReportsEverySectionInDocumentOrder`, `WordSectionInteropTests.OpenedDocument_WithTwoSections_EditsOnlyTheRequestedOne` |
| Crear una sección (`WordDocument.AddSection`) | Probado | `WordSectionTests.AddSection_*`, `WordSectionInteropTests.ProducedDocument_WithTwoSections_IsSchemaValidAndKeepsTheFinalSectionLast` |
| Sección por defecto de un documento nuevo (US Letter, vertical, márgenes de una pulgada) | Probado | `WordSectionTests.Create_HasOneLetterPortraitSectionWithOneInchMargins` |
| Tamaño y orientación de página (`Section.PageSize`, `PageSize`) | Probado | `WordSectionTests.PageSize_*`, `WordSectionInteropTests.ProducedDocument_WithPageSetup_IsSchemaValidAndReadable` |
| Márgenes de página (`Section.Margins`, `PageMargins`) | Probado | `WordSectionTests.Margins_*`, `WordSectionInteropTests.OpenedDocument_WithUnmodelledSectionMarkup_KeepsItWhenPageSetupIsEdited` |
| Tipo de inicio de sección (`Section.StartType`, `w:type`) | Probado | `WordSectionTests.StartType_*`, `WordSectionInteropTests.ProducedDocument_WithPageSetup_IsSchemaValidAndReadable` |
| Orden de esquema de `w:sectPr` y conservación del `w:sectPr` final como último hijo del cuerpo | Probado | `WordSectionInteropTests.ProducedDocument_WithPageSetup_IsSchemaValidAndReadable`, `...WithTwoSections_IsSchemaValidAndKeepsTheFinalSectionLast` |
| Preservar el marcado de sección no modelado (columnas, `w:docGrid`, `w:header`/`w:footer`/`w:gutter` de `w:pgMar`) | Probado | `WordSectionInteropTests.OpenedDocument_WithUnmodelledSectionMarkup_KeepsItWhenPageSetupIsEdited` |
| Archivo ilegible: `w:pgSz`/`w:pgMar`/`w:type` con valor ausente o mal formado | Probado | `WordSectionTests.Margins_ReportAMissingRequiredMargin`, `WordSectionInteropTests.OpenedDocument_WithMalformed*` |
| Columnas, bordes de página, numeración de líneas/páginas, cuadrícula de documento y resto de `w:sectPr` | Limitación conocida | No se modelan; se preservan (ADR 0016) |
| Borrar, mover o fusionar secciones | Limitación conocida | Solo se añaden al final con `AddSection`; quitar un `w:sectPr` fusionaría secciones y queda fuera (ADR 0016) |
| Encabezados y pies por sección y por tipo (`Section.Header`/`Footer`, `HeaderFooterType`) | Probado | `WordHeaderFooterTests.Create_HasNoHeadersOrFooters`, `...EveryType_CanBeAddedAndReadBack`, `...AddHeader_OnDifferentSections_KeepsThemIndependent`, `WordHeaderFooterInteropTests.ProducedDocument_WithEveryHeaderAndFooterType_IsSchemaValidWithReferencesInOrder` |
| Crear un encabezado o pie (`Section.AddHeader`/`AddFooter`) | Probado | `WordHeaderFooterTests.AddHeader_CreatesAnEmptyHeaderAndWritesAReference`, `...AddHeader_IsIdempotent` |
| Eliminar la referencia conservando la parte (`Section.RemoveHeader`/`RemoveFooter`) | Probado | `WordHeaderFooterTests.RemoveHeader_RemovesTheReferenceAndKeepsThePart`, `WordHeaderFooterInteropTests.OpenedDocument_RemovingOneReference_KeepsTheSharedPartForTheOtherSection` |
| Contenido del encabezado o pie: párrafos, tablas y texto (`HeaderFooter`) | Probado | `WordHeaderFooterTests.Header_CanBeEditedThroughItsParagraphs`, `...Footer_CanHoldATable` |
| Compartir una misma parte entre secciones (vínculo sección-parte) | Probado | `WordHeaderFooterTests.AddSection_NewSectionSharesTheHeaderOfThePreviousOne`, `WordHeaderFooterInteropTests.OpenedDocument_WithTwoSectionsSharingAHeader_EditsTheSharedPart` |
| Primera página distinta (`Section.DifferentFirstPage`, `w:titlePg`) | Probado | `WordHeaderFooterTests.DifferentFirstPage_RoundTripsAndClears` |
| Preservar partes huérfanas y contenido no modelado (`w:sdt`) de un encabezado o pie | Probado | `WordHeaderFooterInteropTests.OpenedDocument_WithAnOrphanHeaderPart_PreservesIt`, `...WithUnmodelledHeaderContent_PreservesItWhenTheTextIsEdited` |
| Archivo ilegible: referencia de encabezado/pie sin `w:type`, a una parte ausente o de tipo equivocado, o con `w:titlePg`/tipo mal formado | Probado | `WordHeaderFooterTests.Header_Reports*`, `...DifferentFirstPage_ReportsAMalformedValue`, `WordHeaderFooterInteropTests.OpenedDocument_WithAHeaderReferenceToAFooterPart_Throws` |
| `w:evenAndOddHeaders` (encabezados/pies de páginas pares en todo el documento) | Limitación conocida | Se preserva pero no se modela; sin él Word no muestra el tipo `EvenPage` (ADR 0017) |
| Numeración de páginas y campos OOXML | Probado | `WordFieldTests.AddPageNumber_WritesTheFieldInTheFooter`, `...AddField_WithoutAResult_WritesOnlyTheFieldCharacters`, `...AddField_IsNotCalculated` |
| Cálculo o actualización de los campos | Limitación conocida | Los campos se escriben, no se evalúan; Word los recalcula al abrir o imprimir (ADR 0037) |
| Imágenes inline incrustadas: añadir desde archivo o stream (`Paragraph.AddImage`, `ImageFormat` con PNG, JPEG, GIF, BMP, TIFF, EMF y WMF) | Probado | `WordImageTests.AddImage_FromFile_RoundTripsThroughAFile`, `...AddImage_FromStream_ConvertsPointsToEmus`, `...AddImage_WithEachSupportedFormat_WritesThatContentType`, `WordImageInteropTests.ProducedDocument_WithImage_IsSchemaValidAndRelationshipResolves` |
| Lectura de imágenes inline (`Paragraph.Images`, `Image`) con tamaño, tipo de contenido y bytes | Probado | `WordImageTests.Images_ReadsAnExistingInlineImage`, `...AddImage_FromFile_RoundTripsThroughAFile` |
| Imágenes en párrafos de celdas de tabla y de encabezados/pies | Probado | `WordImageTests.AddImage_InATableCell_RoundTrips`, `...AddImage_InAHeader_RoundTrips`, `WordImageInteropTests.ProducedDocument_WithImagesInBodyTableHeaderAndFooter_IsSchemaValid` |
| Identificadores de dibujo únicos al añadir varias imágenes (`wp:docPr/@id`) | Probado | `WordImageTests.AddImage_Twice_UsesDistinctDrawingIds`, `WordImageInteropTests.ProducedDocument_WithImagesInBodyTableHeaderAndFooter_IsSchemaValid` |
| Preservar relaciones no modeladas al editar (hipervínculo externo y relación de imagen ajena) | Probado | `WordImageInteropTests.OpenedDocument_WithHyperlinkRelationship_PreservesItWhenAnImageIsAdded`, `...OpenedDocument_WithInlineImage_PreservesItWhenTheBodyIsEdited` |
| Archivo ilegible: imagen inline con una relación ausente, sin `wp:extent` o con una relación que no apunta a una parte de imagen | Probado | `WordImageTests.Images_ReportsAMissingImagePart`, `...Images_ReportsAnInlineImageWithoutExtent`, `...Images_ReportsAnImageRelationshipToANonImagePart` |
| Imágenes flotantes (`wp:anchor`), vinculadas (`r:link`), gráficos y otros DrawingML | Limitación conocida | Se preservan pero no se leen ni se crean (ADR 0018); `Paragraph.Images` solo devuelve `wp:inline` con `pic:pic` incrustado |
| Texto alternativo, tamaño automático desde la imagen, deduplicación, reemplazo, eliminación o movimiento de imágenes | Limitación conocida | Fuera de alcance; el tamaño se indica en puntos y cada `AddImage` crea su propia parte (ADR 0018) |
| Imágenes SVG, JPEG 2000, PCX o iconos al añadir | Limitación conocida | No se ofrecen en `ImageFormat`; si ya existen en el documento, se preservan (ADR 0018) |
| Hipervínculos externos: crear (`Paragraph.AddHyperlink`), leer (`Paragraph.Hyperlinks`, `Hyperlink`) y editar texto y destino | Probado | `WordHyperlinkTests.AddHyperlink_RoundTripsThroughAFile`, `...Hyperlink_Text_IsEditableAndKeepsItsDestination`, `...Hyperlink_Destination_IsEditableAndRemovesTheOldRelationship`, `WordHyperlinkInteropTests.ProducedDocument_WithHyperlink_IsSchemaValidAndRelationshipResolves` |
| Hipervínculos en cuerpo, celdas de tabla y encabezados/pies | Probado | `WordHyperlinkTests.AddHyperlink_InATableCell_RoundTrips`, `...AddHyperlink_InAHeaderAndFooter_RoundTrips`, `WordHyperlinkInteropTests.ProducedDocument_WithHyperlinksInBodyTableHeaderAndFooter_IsSchemaValid`, `...OpenedDocument_WithHyperlinkInAHeader_EditsTheRelationshipOfTheHeaderPart` |
| Relación de hipervínculo compartida por varios enlaces: editar uno no cambia los demás | Probado | `WordHyperlinkTests.Hyperlink_Destination_WhenTheRelationshipIsShared_LeavesTheOtherLinkAlone`, `WordHyperlinkInteropTests.OpenedDocument_WithTwoHyperlinksSharingARelationship_KeepsTheRelationshipInUse` |
| Editar el destino elimina la relación anterior cuando ya no la usa ningún enlace de la parte | Probado | `WordHyperlinkTests.Hyperlink_Destination_IsEditableAndRemovesTheOldRelationship`, `WordHyperlinkInteropTests.OpenedDocument_WithHyperlinkInAHeader_EditsTheRelationshipOfTheHeaderPart` |
| Archivo ilegible: hipervínculo con una relación ausente, que no es de hipervínculo o que no es externa | Probado | `WordHyperlinkTests.Hyperlinks_ReportsAMissingRelationship`, `...Hyperlinks_ReportsARelationshipThatIsNotAHyperlink`, `...Hyperlinks_ReportsANonExternalHyperlinkRelationship` |
| Hipervínculos internos a marcadores (`w:anchor`) | Limitación conocida | Válidos y preservados, pero no se crean, leen ni editan; `Paragraph.Hyperlinks` solo devuelve enlaces externos (`WordHyperlinkTests.Hyperlinks_DoesNotReturnAnInternalAnchorLink`, `WordHyperlinkInteropTests.OpenedDocument_WithInternalHyperlink_PreservesItWhenTheBodyIsEdited`, ADR 0019) |
| Atributos `w:tooltip`, `w:history`, `w:tgtFrame` y `w:docLocation` de un hipervínculo | Limitación conocida | Se preservan, pero no se modelan ni se editan (ADR 0019) |
| Estilo de carácter del texto de un hipervínculo | Limitación conocida | `AddHyperlink` no aplica estilo ni formato directo; el llamante puede usar `Hyperlink.Runs` (ADR 0019) |
| Propiedades básicas del documento (`docProps/core.xml`): título, asunto, autor, palabras clave, descripción, último editor, revisión, fechas, categoría, estado, identificador, idioma y versión | Probado | `WordPropertiesTests.Properties_AllCoreValues_RoundTripThroughAFile`, `...Properties_ExistingValues_AreEditableAndRoundTripThroughAStream`, `WordPropertiesInteropTests.ProducedDocument_WithCoreProperties_IsValidAndReadableByTheSdk`, `...ProducedDocument_CoreProperties_UseTheExpectedElementsAndNamespaces` |
| El core part se crea en `docProps/core.xml` solo al asignar un valor; leer un documento sin él no lo crea | Probado | `WordPropertiesTests.Create_HasNoCorePropertiesAndNoCorePart`, `...Properties_ClearingMissingValues_DoesNotCreateTheCorePart`, `WordPropertiesInteropTests.OpenedDocument_WithoutCoreProperties_CreatesTheConventionalCorePart` |
| Propiedades extendidas (`docProps/app.xml`) y personalizadas (`docProps/custom.xml`) | Limitación conocida | Se preservan, pero no se leen ni se editan; `WordPropertiesTests.Properties_PreserveExtendedAndCustomProperties`, `WordPropertiesInteropTests.OpenedDocument_WithCoreProperties_EditsThemAndPreservesUnmodelledParts` (ADR 0020) |
| Fechas de core properties | Limitación conocida | Se guarda el instante; el SDK devuelve `DateTime` local y puede cambiar el offset textual. Comparar con `ToUniversalTime` (ADR 0020) |
| Archivo ilegible: `docProps/core.xml` con una fecha mal formada | Probado | `WordPropertiesTests.Properties_ReportAMalformedCorePartWithATypedError` lanza `InvalidOfficeFileException` en vez de una excepción del SDK |
| Valores de texto con caracteres que XML no puede almacenar | Probado | Se rechazan con `ArgumentException` antes de crear o modificar la parte; tabuladores y saltos de línea se permiten (`WordPropertiesTests.Properties_RejectCharactersXmlCannotStore`, `...Properties_AllowTabsAndLineBreaks`) |
| Búsqueda y reemplazo | Probado | `WordTextSearchTests.FindText_FindsEveryOccurrenceInTheBody`, `...FindText_WithTextSplitAcrossRuns_FindsIt`, `...ReplaceText_ReplacesEveryOccurrenceAndKeepsTheFormatOfTheFirstRun` |
| Búsqueda y reemplazo: encabezados, pies, notas y tablas anidadas; expresiones regulares | Limitación conocida | El alcance es el cuerpo y la búsqueda no cruza párrafos (ADR 0036) |
| Insertar/eliminar/mover elementos | En progreso | Solo añadir párrafos al final del cuerpo y runs, imágenes inline o hipervínculos al final de un párrafo; no se modelan borrar/mover ni insertar en medio |
| Salto de línea y de página | Probado | `w:tab`, `w:br` (`textWrapping`, `page`, `column`) y `w:cr` ajeno: `WordRunContentTests`, `WordTabAndBreakInteropTests` |

## Excel — XLSX (OfficeEngine.Excel)

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Crear libro (con `Sheet1`) | Probado | `ExcelWorkbookTests.Create_HasSingleDefaultWorksheet` |
| Abrir desde archivo/stream | Probado | `ExcelWorkbookTests.Save_ToFile_ThenOpen_RoundTrips` |
| Guardar a archivo (atómico) / stream | Probado | idem y `SaveBytes` |
| Hojas: añadir | Probado | `Worksheets_Add_AppendsAndIsAddressableByName` |
| Hojas: eliminar | Probado | `Worksheets_Remove_DeletesSheetButNotTheLastOne` |
| Hojas: renombrar | Probado | `Worksheets_Rename_ValidatesUniquenessButAllowsCaseChange` |
| Hojas: reordenar | No iniciado | |
| Validación de nombres de hoja | Probado | `Worksheets_Add_RejectsDuplicateAndInvalidNames` |
| Celdas por dirección A1 | Probado | `ExcelWorkbookTests` (varios) |
| Celdas por índice de fila/columna (1-based) | Probado | `ExcelCellIndexTests.CellIndexer_*`, `...CellHandle_ReportsTheRowAndColumnOfItsReference`, `ExcelInteropTests.ProducedPackage_WithCellWrittenByIndex_IsValidAndReadableByTheSdk` |
| Dirección de una celda (`Cell.Row`/`Cell.Column`) | Probado | idem |
| Índice de fila/columna fuera de rango | Probado | `ExcelCellIndexTests.CellIndexer_OutOfRange_*` |
| Tipo: texto | Probado | `Create_ThenEdit_ThenSave_Reopen_RoundTripsValuesAndTypes` |
| Tipo: número | Probado | idem + `Cell_WritesNumbersWithInvariantDecimalSeparator` |
| Tipo: booleano | Probado | idem |
| Tipo: vacío (`null`) | Probado | `Cell_AssignNull_ClearsExistingValue`, `Cell_AssignNull_DoesNotCreateAMissingCell`, `ExcelCellRangeTests.SetValues_ClearsExistingCellsWithoutCreatingMissingOnes` |
| Tipo: fecha/hora (serial + formato) | Probado | `Cell.SetDateTime`, `Cell.Value` con un formato de fecha/hora ya presente, round-trip 1900/1904, seriales no representables devueltos como número; `ExcelInteropTests.ProducedPackage_WithDateAndNumberFormats_IsValidAndReadableByTheSdk` |
| Fecha/hora en rangos | Limitación conocida | `CellRange.GetValues` devuelve `DateTime`; `CellRange.SetValues` rechaza `DateTime` porque cada celda necesita un formato explícito (`ExcelCellRangeTests.GetValues_ReturnsDateTimeForDateFormattedCells`, `...SetValues_DateTime_LeavesTheRangeUnchanged`) |
| Tipo: error | En progreso | Lectura devuelve el texto (`#DIV/0!`); escritura no implementada |
| Shared strings (lectura) | Probado | `ExcelInteropTests.OpenedWorkbook_WithSharedStringsAndMergedCells_*` |
| Fórmula: escribir fórmula (`Cell.Formula`) | Probado | `ExcelFormulaTests.Formula_WriteWithoutCachedValue_RoundTrips`, `...Formula_WithoutLeadingEquals_IsNormalized`, `ExcelInteropTests.OpenedWorkbook_WithFormula_ReturnsCachedValueAndPreservesFormula` |
| Fórmula: leer valor calculado almacenado | Probado | `ExcelInteropTests.OpenedWorkbook_WithFormula_ReturnsCachedValueAndPreservesFormula`, `ExcelFormulaTests.Formula_Null_RemovesFormulaAndCachedValue` |
| Fórmula: resultado almacenado al escribir | Probado | `ExcelFormulaTests.Formula_Set_ClearsStaleCachedValue` (al escribir se declara ausente), `...Formula_Null_RemovesFormulaAndCachedValue` |
| Fórmula: pedir recálculo al abrir (`calcPr/@fullCalcOnLoad`) | Probado | `ExcelFormulaTests.Formula_RequestsRecalculationAfterExistingDefinedNames` |
| Fórmula: cálculo (motor propio) | Probado | `FormulaEngineTests` (105 casos) y `FormulaDifferentialTests` (31 fórmulas contra un motor externo); matriz en `docs/formula-matrix.md` (ADR 0043) |
| Fórmula: nombres de las funciones en español | Probado | `FormulaEngineTests.Evaluate_WithTheSpanishNameOfAFunction_ComputesTheSameValue` (29 funciones), `FormulaNameTests` (traducción al nombre canónico) y `ExcelFormulaTests.Formula_WithTheSpanishNameOfAFunction_IsStoredWithTheCanonicalName` (ADR 0043) |
| Fórmula: funciones volátiles, matrices, referencias externas y nombres definidos | Limitación conocida | Fuera de la matriz publicada: devuelven `#NAME?` o `#VALUE!`, nunca un valor inventado (ADR 0043) |
| Fórmulas compartidas, de matriz y de tabla | Limitación conocida | Se conservan al leer y guardar, pero no se escriben ni se expanden; una compartida sin texto lee `null` (ADR 0025) |
| Fórmula: validación semántica de la expresión | Limitación conocida | No se valida: solo se traducen los nombres de las funciones al nombre canónico (que es lo que Excel guarda); la calcula el motor que la abre (ADR 0025, ADR 0043) |
| Rangos (`A1:B2`) | Probado | `ExcelCellRangeTests.GetRange_*`, `...Cells_*`, `...GetValues_*`, `...SetValues_*`, `ExcelInteropTests.ProducedPackage_WithRangeValues_IsValidAndReadableByTheSdk` |
| Rangos: recorrido en orden de fila | Probado | `ExcelCellRangeTests.Cells_EnumeratesInRowMajorOrder`, `...Cells_IsALiveView` |
| Rangos: lectura/escritura masiva (`GetValues`/`SetValues`) | Probado | `ExcelCellRangeTests.GetValues_ReadsCellsInRowMajorOrderAndReturnsNullForMissingOnes`, `...SetValues_WritesEveryCellAndGetValuesReadsItBack`, `...RangeValues_RoundTripThroughSaveAndOpen` |
| Rangos: celdas ausentes no se crean al leer ni al escribir `null` | Probado | `ExcelCellRangeTests.Cells_ReadingDoesNotCreateMissingCells`, `...SetValues_ClearsExistingCellsWithoutCreatingMissingOnes` |
| Rangos: referencia inválida o invertida | Probado | `ExcelCellRangeTests.GetRange_InvalidReference_ThrowsArgumentException`, `...GetRange_IntOverload_ReversedCorners_ThrowsArgumentException` |
| Rangos: dimensiones de la matriz de valores | Probado | `ExcelCellRangeTests.SetValues_RequiresExactDimensions` |
| Rangos: tipo no soportado sin escritura parcial | Probado | `ExcelCellRangeTests.SetValues_UnsupportedValue_LeavesTheRangeUnchanged` |
| Filas/columnas: copiar/mover | No iniciado | |
| Estilos reutilizables (`Cell.Style`): fuente, relleno sólido, cuatro bordes y alineación | Probado | `ExcelCellStyleTests.Style_*`, `ExcelInteropTests.ProducedPackage_WithCellStyles_IsValidAndReadableByTheSdk` |
| Estilos: reutilización de `fonts`/`fills`/`borders`/`cellXfs` | Probado | `ExcelCellStyleTests.Style_SharedAcrossCells_ReusesCellFormatsAndComponents`, `...Style_DifferentStyles_UseDifferentCellFormats` |
| Estilos: reset con un estilo vacío y formato numérico conservado | Probado | `ExcelCellStyleTests.Style_DefaultResetsModeledFormattingAndKeepsNumberFormat`, `...Style_EmptyBordersOnMissingCell_DoesNotCreateCell` |
| Estilos: propiedades no modeladas de fuente, borde y alineación | Probado | `ExcelCellStyleTests.Style_PreservesUnmodelledFontBorderAndAlignmentProperties` |
| Estilos: rellenos con patrón, degradados o color de tema; protección, rotación, sangría y resto de alineación | Limitación conocida | No se modelan; el relleno no sólido se reemplaza al aplicar un estilo y el resto se preserva (ADR 0024) |
| Estilos con nombre, estilos de tabla, formato condicional y `dxf` | Limitación conocida | No se modelan; la herencia `cellStyleXfs`/`xfId` se conserva sin resolver (ADR 0024) |
| Formatos numéricos (`Cell.NumberFormat`) | Probado | `ExcelDateAndNumberFormatTests.NumberFormat_*`, `...OpenedWorkbook_WithBuiltInDateFormat_ReadsDateTimeAndFormat`, `...WithCustomDateFormat_ReadsDateTime` |
| Formatos numéricos integrados no modelados | Limitación conocida | Se preservan; `NumberFormat` devuelve `null` para ids no modelados (por ejemplo 23–26 y 41–44) y para códigos dependientes de la configuración regional (27–36 y 50–58) |
| Rangos con nombre | Probado | `ExcelDefinedNameTests.AddNamedRange_CreatesWorkbookScopedName`, `...AddLocalNamedRange_CreatesSheetScopedName`, `...DefinedNames_ReadsTheNamesOfAnotherWorkbook`, `...RemoveSheet_MovesTheScopeOfTheRemainingLocalNames` |
| Rangos con nombre: referencias 3D, externas, nombres dinámicos y fórmulas con nombre | Limitación conocida | Se conservan y se listan, pero no se interpretan (`IsRange` es `false`) (ADR 0026) |
| Validación de datos | Probado | `ExcelDataValidationTests.AddDataValidation_WithAList_WritesTheRule`, `...AddDataValidation_WithANumber_WritesTheOperator`, `...DataValidations_ReadsTheRulesOfAnotherWorkbook` |
| Validación de datos: reglas de fecha, hora, longitud y fórmula | Limitación conocida | Se leen y se conservan, pero no se pueden crear (ADR 0027) |
| Autofiltros | Probado | `ExcelAutoFilterTests.AutoFilter_Set_WritesTheRange`, `...AutoFilter_ReadsTheFilterOfAnotherWorkbook` |
| Autofiltros: criterios de filtrado y `sortState` | Limitación conocida | Se conservan al guardar la hoja intacta y se descartan al escribir el filtro (ADR 0028) |
| Tablas estructuradas | Probado | `ExcelTableTests.AddTable_CreatesTheTablePartAndTheReference`, `...Tables_ReadsTheTablesOfAnotherWorkbook`, `...RemoveTable_RemovesTheTablePart` |
| Tablas estructuradas: totales, columnas calculadas y estilo | Limitación conocida | Se conservan al guardar; el estilo se fija en `TableStyleMedium2` (ADR 0029) |
| Congelar paneles | Probado | `ExcelSheetLayoutTests.FreezePanes_Set_WritesThePane`, `...FreezePanes_ReadsThePaneOfAnotherWorkbook` |
| Ocultar filas/columnas, alto/ancho | Probado | `ExcelSheetLayoutTests.SetRowHeight_WritesTheHeight`, `...SetRowHidden_HidesAndShowsTheRow`, `...SetColumnWidth_WritesTheWidth`, `...SetColumnHidden_HidesAndShowsTheColumn` |
| Paneles divididos, zoom y selecciones visibles | Limitación conocida | Se conservan; solo se modelan los paneles congelados (ADR 0030) |
| Configuración de impresión y página | Probado | `ExcelPrintSettingsTests.PageMargins_Set_WritesTheMargins`, `...PageSetup_Set_WritesTheSetup`, `...PrintOptions_Set_WritesTheOptions` |
| Configuración de impresión: encabezado y pie, área de impresión | Limitación conocida | Se conservan; no hay API (ADR 0031) |
| Hipervínculos | Probado | `ExcelHyperlinkTests.AddHyperlink_WritesTheAddressAndTheRelationship`, `...AddInternalHyperlink_WritesTheLocation`, `...RemoveHyperlink_RemovesTheLinkAndTheRelationship` |
| Notas/comentarios | Probado | `ExcelCommentTests.AddComment_WritesTheCommentPartAndTheDrawing`, `...Comments_ReadsTheNotesOfAnotherWorkbook` |
| Notas: comentarios con hilo y tamaño del cuadro | Limitación conocida | Se conservan; la posición del cuadro de la nota es aproximada (ADR 0032) |
| Imágenes | Probado | `ExcelImageTests.AddImage_WritesTheMediaPartAndTheDrawing`, `...Images_ReadsTheImagesOfAnotherWorkbook`, `...RemoveImage_RemovesTheAnchor` |
| Imágenes: anclajes de dos celdas, giros y recortes | Limitación conocida | Se conservan; solo se crean anclajes de una celda (ADR 0033) |
| Gráficos | Probado | `ExcelChartTests.AddChart_WritesTheChartPartAndTheDrawing`, `...AddChart_WithEveryType_WritesAValidPackage`, `...Charts_ReadsTheChartsOfAnotherWorkbook` |
| Gráficos: varias series, tipos de eje y áreas de combinación | Limitación conocida | Se conservan; los gráficos que crea esta versión tienen una serie y son de columna, barra, línea o sector (ADR 0034) |
| Escritura de cadenas como *inline strings* | No iniciado | El texto se escribe en la tabla de cadenas compartidas (`xl/sharedStrings.xml`), con deduplicación (ADR 0035) |
| Tabla de cadenas compartidas (escritura) | Probado | `ExcelSharedStringTests.Strings_AreWrittenToTheSharedStringTable`, `...Strings_RoundTripThroughSaveAndOpen`, `ExcelInteropTests.ProducedPackage_ContainsSharedStringsAndIsReadableByTheSdk` |
| Conservación de contenido no modelado | Probado | `OpenedWorkbook_WithSharedStringsAndMergedCells_*` (mergeCells preservado) |
| Diagnóstico de contenido no modelado | Probado | idem |
| Detección de fila sin índice (archivo malformado) | Probado | `OpenedWorkbook_WithRowMissingIndex_ThrowsInvalidOfficeFileExceptionInsteadOfReturningNull` |
| Validación del paquete producido (SDK externo) | Probado | `ExcelInteropTests.ProducedPackage_IsAValidSpreadsheetDocument`, `...ProducedPackage_WithCellStyles_IsValidAndReadableByTheSdk` |
| Exportación CSV | No iniciado | Auxiliar y opcional |

## PDF (OfficeEngine.Pdf)

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Interfaz de conversión | Probado | `OfficeEngine.Pdf`: `IPdfConverter`, `PdfSourceFormat`, `PdfConversionOptions`, `PdfConversionResult`, `PdfConversionException` (ADR 0038) |
| Adaptador LibreOffice headless | Probado | `LibreOfficePdfConverterTests` (integración); LibreOffice 26.8.1.1 verificado el 2026-10-21, MPL-2.0 (ADR 0038) |
| DOCX→PDF / XLSX→PDF | Probado | `LibreOfficePdfConverterTests.Convert_WordDocument_ProducesAPdf`, `...Convert_ExcelWorkbook_ProducesAPdf` |
| Exportación desde el estado en memoria | Probado | `LibreOfficePdfConverterTests.Convert_UsesTheBytesItReceives` |
| Aislamiento del motor (perfil, temporales, tiempo de espera, límites) | Probado | `...Convert_LeavesNoTemporaryDirectoryBehind`, `...Convert_WithATooShortTimeout_Fails`, `...Convert_WithAnInputLargerThanTheLimit_Throws` |
| Motor ausente o versión anterior a la mínima | Probado | `...Convert_WithAMissingEngine_Fails`, `TheEngine_ReportsItsPathAndVersion` |
| Equivalencia visual con Microsoft Office | Limitación conocida | El PDF depende de la versión de LibreOffice, de las fuentes y de la configuración del sistema (ADR 0038) |
| Opciones del PDF (páginas, marcas de agua, cifrado, metadatos) | No iniciado | El adaptador solo produce el PDF que el motor genera |

## DOC/XLS heredados (OfficeEngine.Legacy)

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Convertir `.doc` a `.docx` y `.xls` a `.xlsx` | Probado | `LegacyConverterTests.Convert_WordDocument_ProducesAModernDocument`, `...Convert_ExcelWorkbook_ProducesAModernWorkbook` (ADR 0042) |
| Leer un `.doc` o un `.xls` directamente | No iniciado | Solo se convierte: el resultado se abre con la API moderna (ADR 0042) |
| Editar un archivo heredado en su formato | No iniciado | La edición y el guardado en el formato heredado no existen (ADR 0042) |
| Guardar en el formato heredado | No iniciado | Guardar escribe el formato moderno; el original no se conserva (ADR 0042) |
| Pérdidas de la conversión (macros, OLE, maquetación, metadatos) | Limitación conocida | Dependen del motor y del archivo; se documentan de forma general, no caso por caso (ADR 0042) |

## Seguridad

| Capacidad | Estado | Evidencia / notas |
| --- | --- | --- |
| Límite de tamaño de archivo | Probado | `ResourceLimitTests.*_ExceedingMaxFileSize_*` |
| Límite de número de entradas | Probado | `Word_Open_ExceedingMaxEntryCount_ThrowsResourceLimit` |
| Límite de tamaño total descomprimido | Probado | `Word_Open_ExceedingTotalUncompressedSize_ThrowsResourceLimit` |
| Límite de profundidad de XML | Probado | `HostilePackageTests.Word_WithADeeplyNestedPart_ThrowsATypedError`, `...Excel_WithADeeplyNestedPart_ThrowsATypedError`, `...TheXmlDepthLimit_IsConfigurable` (ADR 0039) |
| DTD y entidades externas | Probado | `HostilePackageTests.Word_WithADocumentTypeDefinition_ThrowsATypedError`, `...Excel_WithAnExternalEntity_ThrowsATypedError` |
| Archivos corruptos: excepciones siempre tipadas | Probado | `CorruptPackageTests.Word_WithMutatedPackages_FailsWithATypedErrorOrOpens`, `...Excel_WithMutatedPackages_...` (ADR 0039) |
| Tiempo máximo de apertura y cancelación | No iniciado | Exige el modelo asíncrono (ADR 0004); la regla para adoptarlo está en el ADR 0040 |
| Documentos grandes (10 000 párrafos o filas) | Probado | `LargeDocumentTests.ALargeWordDocument_RoundTrips`, `...ALargeExcelWorkbook_RoundTrips` (ADR 0040) |
| Rendimiento medido y reproducible | Probado | `OfficeEngine.Benchmarks` mide seis escenarios y los compara con `benchmarks/baseline.txt`; CI publica el informe (ADR 0041) |
| Regresión de round-trip sobre documentos generados | Probado | `WordRoundTripRegressionTests.GeneratedDocuments_RoundTripWithoutLosingText`, `ExcelRoundTripRegressionTests.AGeneratedWorkbook_RoundTripsWithoutLosingValues` |
| Coste lineal al escribir muchas celdas | Probado | `ExcelPerformanceTests.WritingManyCells_DoesNotAllocateProportionallyToTheSquare` (ADR 0041) |
| Detección de bombas ZIP por ratio | Probado | `*_CompressionBomb_ThrowsResourceLimitWithDefaultLimits` |
| Rechazo de entradas que no son ZIP | Probado | `*_InputThatIsNotAZip_*` |
| Sin extracción a rutas arbitrarias | Probado (por diseño) | No existe ninguna API de extracción |
| Sin ejecución de macros/OLE/scripts | Probado (por diseño) | No existe ninguna API de ejecución |
| Sin `eval` de fórmulas | Probado (por diseño) | Las fórmulas se escriben y se leen como texto; solo se lee el valor en caché y el recálculo se delega en el motor que abre el archivo |
| Explicitud ante filas malformadas | Probado | `OpenedWorkbook_WithRowMissingIndex_*` |
| Timeout / cancelación de operaciones | No iniciado | Requiere almacén con cancelación real (ADR 0004) |
| Cifrado con contraseña | No iniciado | Un archivo cifrado no es un ZIP y se rechaza con `InvalidOfficeFileException`; sin soporte de contraseñas |
| Protección de hoja/libro | No iniciado | No se interpreta ni se desprotege. Por el diseño de preservación no debería eliminarse, pero **no está probado específicamente** aún |

## Limitaciones globales conocidas (Fase 1, 2.1, 2.2, 2.2b y 2.3)

1. El documento completo se mantiene en memoria; el tamaño máximo lo fija
   `OfficeEngineLimits.MaxFileSizeBytes` (256 MiB por defecto).
2. Toda la I/O es síncrona (ADR 0004).
3. Se modelan el texto y el formato de carácter de los runs (familia y tamaño de fuente, negrita,
   cursiva, subrayado y color RGB), el formato de párrafo (alineación, sangría, espaciado y
   tabulaciones), el contenido de run tabulación y saltos de línea/página, las **referencias de
   estilo** (`w:pStyle`/`w:rStyle`), el **catálogo de estilos de solo lectura** con su herencia
   `w:basedOn`, las **referencias de lista** (`w:numPr`), el **catálogo de listas de solo lectura** y
   la **creación de listas de un solo nivel**, las tablas (estructura, contenido, propiedades y
   combinación), las secciones y sus encabezados/pies, las **imágenes inline incrustadas**, los
   **hipervínculos externos** y las **propiedades básicas del documento** (`docProps/core.xml`). En
   Excel se modelan celdas, rangos, fechas, formatos numéricos y estilos reutilizables (fuente,
   relleno sólido, cuatro bordes y alineación).
   **No** se modelan el resto de atributos de `w:rPr` (resaltado,
   tachado, sombreados, colores de tema…), la creación/edición de definiciones de estilo, la
   personalización de niveles de lista (texto, inicio, reinicios), las imágenes flotantes o
   vinculadas, los gráficos, los hipervínculos internos a marcadores ni los atributos de hipervínculo
   distintos del destino (`w:tooltip`, `w:history`, `w:tgtFrame`, `w:docLocation`), ni las propiedades
   extendidas (`docProps/app.xml`) y personalizadas (`docProps/custom.xml`); ese contenido
   **se preserva** en Word y Excel y, cuando es detectable, se reporta por diagnósticos.
4. Los diagnósticos se calculan al abrir y no se recalculan tras editar.
5. No hay cálculo de fórmulas ni de campos OOXML: los campos de Word se escriben con el resultado
   almacenado que se les da y Word los recalcula al abrir o al imprimir.
6. No hay soporte de DOC/XLS. El PDF se produce con un motor externo (LibreOffice headless), sin
   equivalencia visual garantizada con Microsoft Office.
7. No hay soporte de documentos cifrados.
8. Cuando un valor almacenado no se puede leer o mapear, la lectura falla con
   `InvalidOfficeFileException`; no se devuelve ningún valor inventado y no se escapa ninguna
   excepción de la BCL (ver `docs/adr/0008-character-formatting-model.md`).
