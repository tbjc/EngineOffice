# Matriz de fórmulas soportadas

Publicada con el incremento 6.2 (ADR 0043). Es el contrato exacto de lo que el motor
`OfficeEngine.Formulas` calcula: lo que no está aquí, **no** se calcula.

Son **32 funciones**, con **61 nombres**: el nombre en inglés y, cuando existe, el de Excel en
español.

## Los nombres en español

El motor acepta el nombre de cada función **en inglés y en español** (`=SUM(A1:A2)` y
`=SUMA(A1:A2)` son la misma fórmula), porque las hojas se escriben en español.

Excel guarda los nombres de las funciones **en inglés** dentro del archivo, y lo único que cambia
es lo que muestra la interfaz. Por eso, cuando la biblioteca escribe una fórmula en un libro, la
traduce al nombre canónico: `=SUMA(A1:A2)` se guarda como `=SUM(A1:A2)`, que es lo que Excel
evalúa. El texto de los literales no se toca.

## Cómo se evalúa

| Elemento | Soportado | Notas |
| --- | --- | --- |
| Números | Sí | `2`, `2.5`, `1E3` |
| Textos | Sí | `"hola"`; la comilla doble se escapa con `""` |
| Booleanos | Sí | `TRUE`, `FALSE` |
| Valores de error | Sí | `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!` |
| Referencias a celdas | Sí | `A1`, `$A$1`, `Hoja1!A1`, `'Mi Hoja'!A1` |
| Rangos | Sí | `A1:B2`, como argumento de una función |
| Paréntesis | Sí | `(A1+A2)*2` |
| Operadores | Sí | `+` `-` `*` `/` `^` `%` `&` `=` `<>` `<` `>` `<=` `>=` |
| Nombres definidos | **No** | Devuelve `#NAME?` |
| Referencias externas | **No** | Devuelve `#NAME?` |
| Fórmulas de matriz | **No** | Un rango en una operación aritmética devuelve `#VALUE!` |
| Funciones volátiles | **No** | `HOY`, `AHORA`, `ALEATORIO` y compañía no están |

## Funciones

| Función | Argumentos | Qué hace |
| --- | --- | --- |
| `SUMA` / `SUM` | 1..n | Suma los números. En un rango, ignora el texto y los booleanos. |
| `PRODUCTO` / `PRODUCT` | 1..n | Multiplica los números. |
| `PROMEDIO` / `AVERAGE` | 1..n | Media de los números. Sin números, `#DIV/0!`. |
| `MIN` / `MAX` | 1..n | El menor o el mayor número. Sin números, `0`. |
| `CONTAR` / `COUNT` | 1..n | Cuenta los números. |
| `CONTARA` / `COUNTA` | 1..n | Cuenta las celdas no vacías. |
| `Y` / `AND` | 1..n | Verdadero si todos los argumentos lo son. |
| `O` / `OR` | 1..n | Verdadero si alguno lo es. |
| `NO` / `NOT` | 1 | Invierte un booleano. |
| `SI` / `IF` | 2..3 | Devuelve la segunda o la tercera rama. **Solo se evalúa la rama elegida.** |
| `SI.ERROR` / `IFERROR` | 2 | Devuelve el segundo valor si el primero es un error. |
| `ABS` | 1 | Valor absoluto. |
| `ENTERO` / `INT` | 1 | Parte entera. |
| `RAIZ` / `SQRT` | 1 | Raíz cuadrada. Negativo → `#NUM!`. |
| `REDONDEAR` / `ROUND` | 1..2 | Redondea, con el empate hacia el infinito. |
| `REDONDEAR.MAS` / `ROUNDUP` | 1..2 | Redondea hacia arriba. |
| `REDONDEAR.MENOS` / `ROUNDDOWN` | 1..2 | Redondea hacia abajo. |
| `RESIDUO` / `MOD` | 2 | Resto, con el signo del divisor. Divisor cero → `#DIV/0!`. |
| `POTENCIA` / `POWER` | 2 | Potencia. |
| `CONCATENAR` / `CONCATENATE` | 1..n | Concatena los textos. |
| `LARGO` / `LEN` | 1 | Longitud del texto. |
| `IZQUIERDA` / `LEFT` | 1..2 | Los primeros caracteres. |
| `DERECHA` / `RIGHT` | 1..2 | Los últimos caracteres. |
| `EXTRAE` / `MID` | 3 | Una parte del texto. |
| `MAYUSC` / `UPPER` | 1 | En mayúsculas. |
| `MINUSC` / `LOWER` | 1 | En minúsculas. |
| `ESPACIOS` / `TRIM` | 1 | Quita los espacios sobrantes. |
| `ESBLANCO` / `ISBLANK` | 1 | Verdadero si la celda está vacía. |
| `ESNUMERO` / `ISNUMBER` | 1 | Verdadero si el valor es un número. |
| `ESTEXTO` / `ISTEXT` | 1 | Verdadero si el valor es un texto. |
| `ESERROR` / `ISERROR` | 1 | Verdadero si el valor es un error. |

Los nombres se aceptan en inglés y en español, sin distinguir mayúsculas y minúsculas.

## Reglas de conversión

| Situación | Resultado |
| --- | --- |
| Celda vacía en una operación aritmética | `0` |
| Celda vacía en una concatenación | `""` |
| Texto en una operación aritmética | `#VALUE!`, salvo que se pueda leer como número |
| Texto en una comparación | Se compara sin distinguir mayúsculas, y es **mayor** que cualquier número |
| Booleano en una operación aritmética | `1` o `0` |
| Booleano en una concatenación | `TRUE` o `FALSE` |
| Número en un texto | Formato general, con 15 cifras significativas: `2` → `2`, `2.5` → `2.5`, `1/3` → `0.333333333333333` |
| Un error en cualquier operando | El error se propaga al resultado, salvo que la función lo acepte (`ISERROR`, `IFERROR`) |

## Diferencias conocidas con Excel

- El motor **no sigue las dependencias**: evalúa sobre los **valores** que hay en las celdas, no
  sobre las fórmulas que contienen. Una referencia circular no es un problema, y tampoco se
  recalcula nada.
- Un rango solo es válido como argumento de una función: `=A1:A2+1` devuelve `#VALUE!`.
- No hay `HOY`, `AHORA`, `ALEATORIO` ni ninguna otra función volátil: sus resultados no serían
  reproducibles.
- Las fechas se leen como números: el motor no conoce el sistema de fechas del libro.
- Los errores y los booleanos se escriben en inglés, como en Excel. Un exportador en otro idioma
  puede mostrarlos traducidos: es una diferencia de idioma, no de significado.
