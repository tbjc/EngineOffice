---
id: listas
title: Listas
sidebar_label: Listas
---

# Listas

Una lista se crea con un formato y se aplica a un párrafo con un nivel.

## API

| Miembro | Qué hace |
| --- | --- |
| `WordDocument.AddList(WordListFormat format)` | Crea una lista de un solo nivel y devuelve su definición. |
| `WordDocument.Lists` | Las listas del documento, en solo lectura. |
| `WordDocument.FindList(int id)` | Busca una lista por su identificador. |
| `Paragraph.ApplyList(WordList list, int level = 0)` | Aplica la lista al párrafo. |
| `Paragraph.ListId` | La lista a la que pertenece el párrafo. |
| `Paragraph.ListLevel` | El nivel dentro de la lista. |
| `WordList.Levels` | Los niveles de la lista. |
| `WordListLevel.Format`, `Text`, `Start`, `Justification` | Los datos de un nivel. |

## Ejemplo

```csharp
using OfficeEngine;
using OfficeEngine.Word;

using WordDocument document = WordDocument.Create();

WordList items = document.AddList(WordListFormat.Bullet);

document.AddParagraph("Primer punto").ApplyList(items);
document.AddParagraph("Segundo punto").ApplyList(items);

WordList numbers = document.AddList(WordListFormat.Decimal);
document.AddParagraph("Primero").ApplyList(numbers);
document.AddParagraph("Segundo").ApplyList(numbers);

document.Save("informe.docx");
```

## Notas

- `AddList` crea una lista de **un solo nivel**: personalizar el texto del nivel, el inicio o los
  reinicios no se modela (ADR 0012).
- Los formatos de numeración que OOXML define y la biblioteca no modela (por ejemplo `chicago`)
  lanzan `UnsupportedOfficeFeatureException` en vez de inventarse.