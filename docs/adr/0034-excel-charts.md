# ADR 0034 — Gráficos en la hoja

- Estado: Aceptado
- Fecha: 2026-10-21
- Fase: 3, incremento 3.10b

## Contexto

Un gráfico son cuatro cosas:

1. La parte del gráfico (`xl/charts/chart1.xml`), con el espacio de nombres de gráficos
   (`c:chartSpace`).
2. El **dibujo** de la hoja, que contiene un `xdr:graphicFrame` con un `a:graphicData` que
   apunta a la parte del gráfico.
3. La relación entre el dibujo y la parte del gráfico.
4. Los datos: las categorías y los valores, con su referencia a los rangos de la hoja **y** su copia
   (la caché) dentro del gráfico.

A diferencia de una imagen, un gráfico tiene estructura propia: tipo, serie, ejes, leyenda y
escala. Modelarlo entero (múltiples series, tipos de eje secundarios, etiquetas, líneas de tendencia,
áreas de combinación) es un proyecto mayor que la propia biblioteca.

## Decisión

### 1. API

- `Worksheet.AddChart(cellReference, type, categoriesRange, valuesRange, title, seriesName,
  widthPoints, heightPoints)`: añade un gráfico con **una sola serie**.
- `Worksheet.Charts`: catálogo de solo lectura con los gráficos de la hoja.
- `Worksheet.RemoveChart(name)`: elimina el gráfico, su marco y su parte.
- `ExcelChart`: expone `Name`, `Reference`, `Type`, `CategoriesRange`, `ValuesRange`, `Title` y
  `SeriesName`.
- `ExcelChartType`: `Column`, `Bar`, `Line` y `Pie`.

### 2. Tipos soportados

| Tipo | Elemento | Ejes | Leyenda |
| --- | --- | --- | --- |
| Columna | `c:barChart` con `barDir="col"` | sí | sí |
| Barra | `c:barChart` con `barDir="bar"` | sí | sí |
| Línea | `c:lineChart` | sí | sí |
| Sector | `c:pieChart` | no | no |

Los tipos enumerados son los que esta versión escribe. Un gráfico de otro tipo (dispersión,
área, combo…) se conserva al guardar, pero no se interpreta.

### 3. Los datos van en la referencia y en la caché

- La **referencia** es absoluta y con el nombre de la hoja: `Hoja1!$A$2:$A$5`. Es lo que permite
  al motor que abre el archivo recalcular el gráfico.
- La **caché** copia los valores que hay en el momento de crear el gráfico: los textos de las
  categorías y los números de los valores. Las celdas vacías se dejan como huecos (no se escribe un
  punto) y el `pointCount` declara el tamaño del rango.

La caché es una foto: si los datos cambian después, el gráfico sigue mostrando la foto hasta que
Excel lo recalcula.

### 4. Geometría

El gráfico se ancla a una celda con `xdr:oneCellAnchor` y un `xdr:graphicFrame`, con un tamaño por
defecto de 480 × 288 puntos. Igual que las imágenes, el tamaño no se recalcula con las celdas.

## Consecuencias

- La API pública nueva es `ExcelChart`, `ExcelChartType`, `Worksheet.AddChart`, `Worksheet.Charts` y
  `Worksheet.RemoveChart`.
- El paquete producido puede contener una parte de gráfico, con su relación al dibujo.
- Un gráfico creado con esta biblioteca tiene una sola serie y cuatro tipos posibles.

## Verificación

- `ExcelChartTests`: escritura de la parte de gráfico y del dibujo con round-trip; los cuatro tipos
  producen un paquete válido; varios gráficos; valores inválidos; eliminación; lectura de un
  archivo ajeno; preservación al guardar; libro liberado.
- `OoxmlValidation.CreateXlsxWithChart` construye el fixture con el SDK, con el `chartSpace` escrito
  a mano, y `TestPackage` comprueba el paquete.
