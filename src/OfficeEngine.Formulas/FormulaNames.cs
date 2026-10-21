using System.Text;

namespace OfficeEngine.Formulas;

/// <summary>
/// The names of the functions of the matrix. Excel stores the names of the functions in English, but
/// the sheets this library writes are written in Spanish (ADR 0043), so the Spanish name of a
/// function is an alias of its canonical name.
/// </summary>
internal static class FormulaNames
{
    /// <summary>The Spanish name of a function, mapped to the canonical name.</summary>
    private static readonly Dictionary<string, string> SpanishNames =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUMA"] = "SUM",
            ["PRODUCTO"] = "PRODUCT",
            ["PROMEDIO"] = "AVERAGE",
            ["CONTAR"] = "COUNT",
            ["CONTARA"] = "COUNTA",
            ["Y"] = "AND",
            ["O"] = "OR",
            ["NO"] = "NOT",
            ["SI"] = "IF",
            ["SI.ERROR"] = "IFERROR",
            ["ENTERO"] = "INT",
            ["RAIZ"] = "SQRT",
            ["REDONDEAR"] = "ROUND",
            ["REDONDEAR.MAS"] = "ROUNDUP",
            ["REDONDEAR.MENOS"] = "ROUNDDOWN",
            ["RESIDUO"] = "MOD",
            ["POTENCIA"] = "POWER",
            ["CONCATENAR"] = "CONCATENATE",
            ["LARGO"] = "LEN",
            ["IZQUIERDA"] = "LEFT",
            ["DERECHA"] = "RIGHT",
            ["EXTRAE"] = "MID",
            ["MAYUSC"] = "UPPER",
            ["MINUSC"] = "LOWER",
            ["ESPACIOS"] = "TRIM",
            ["ESBLANCO"] = "ISBLANK",
            ["ESNUMERO"] = "ISNUMBER",
            ["ESTEXTO"] = "ISTEXT",
            ["ESERROR"] = "ISERROR",
        };

    /// <summary>Gets the canonical name of a function, or the name itself when it has no alias.</summary>
    internal static string CanonicalName(string name)
        => SpanishNames.TryGetValue(name, out string? canonical) ? canonical : name;

    /// <summary>Gets the Spanish names of the functions.</summary>
    internal static IEnumerable<string> SpanishAliases => SpanishNames.Keys;

    /// <summary>
    /// Replaces the Spanish names of the functions with the canonical ones. The text of the string
    /// literals is left untouched, and only a name that is a call is translated.
    /// </summary>
    internal static string Canonicalize(string formula)
    {
        StringBuilder builder = new(formula.Length);
        int index = 0;
        while (index < formula.Length)
        {
            char current = formula[index];
            if (current == '"')
            {
                index = CopyLiteral(formula, index, builder);
                continue;
            }

            if (FormulaTokenizer.IsIdentifierStart(current))
            {
                int start = index;
                while (index < formula.Length && FormulaTokenizer.IsIdentifierPart(formula[index]))
                {
                    index++;
                }

                string word = formula[start..index];
                builder.Append(IsCall(formula, index) ? CanonicalName(word) : word);
                continue;
            }

            builder.Append(current);
            index++;
        }

        return builder.ToString();
    }

    /// <summary>Copies a string literal, including the quotes and the escaped quotes.</summary>
    private static int CopyLiteral(string formula, int start, StringBuilder builder)
    {
        int index = start;
        builder.Append(formula[index]);
        index++;
        while (index < formula.Length)
        {
            builder.Append(formula[index]);
            if (formula[index] != '"')
            {
                index++;
                continue;
            }

            if (index + 1 < formula.Length && formula[index + 1] == '"')
            {
                builder.Append(formula[index + 1]);
                index += 2;
                continue;
            }

            return index + 1;
        }

        return index;
    }

    /// <summary>Gets whether the name that ends at <paramref name="index"/> is a call.</summary>
    private static bool IsCall(string formula, int index)
    {
        int next = index;
        while (next < formula.Length && char.IsWhiteSpace(formula[next]))
        {
            next++;
        }

        return next < formula.Length && formula[next] == '(';
    }
}
