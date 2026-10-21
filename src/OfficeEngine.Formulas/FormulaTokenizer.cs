using System.Text;
using System.Globalization;

namespace OfficeEngine.Formulas;

/// <summary>The kinds of token a formula is made of.</summary>
internal enum FormulaTokenKind
{
    Number,
    Text,
    Boolean,
    Error,
    Identifier,
    Operator,
    LeftParenthesis,
    RightParenthesis,
    Comma,
    Colon,
    Bang,
    End,
}

/// <summary>One token of a formula.</summary>
internal sealed class FormulaToken
{
    internal FormulaToken(FormulaTokenKind kind, string text, double number = 0d, bool boolean = false, FormulaError? error = null)
    {
        Kind = kind;
        Text = text;
        Number = number;
        Boolean = boolean;
        Error = error;
    }

    internal FormulaTokenKind Kind { get; }

    internal string Text { get; }

    internal double Number { get; }

    internal bool Boolean { get; }

    internal FormulaError? Error { get; }

    /// <summary>Gets whether the token is the given operator.</summary>
    internal bool IsOperator(string value) => Kind == FormulaTokenKind.Operator && Text == value;
}

/// <summary>Splits a formula into its tokens.</summary>
internal static class FormulaTokenizer
{
    internal static List<FormulaToken> Tokenize(string formula)
    {
        List<FormulaToken> tokens = [];
        int index = 0;
        if (index < formula.Length && formula[index] == '=')
        {
            index++;
        }

        while (index < formula.Length)
        {
            char current = formula[index];
            if (char.IsWhiteSpace(current))
            {
                index++;
                continue;
            }

            if (char.IsAsciiDigit(current) || (current == '.' && index + 1 < formula.Length && char.IsAsciiDigit(formula[index + 1])))
            {
                tokens.Add(ReadNumber(formula, ref index));
                continue;
            }

            if (current == '"')
            {
                tokens.Add(ReadText(formula, ref index));
                continue;
            }

            if (current == '#')
            {
                tokens.Add(ReadError(formula, ref index));
                continue;
            }

            if (current == '\'')
            {
                tokens.Add(new FormulaToken(FormulaTokenKind.Text, ReadQuoted(formula, ref index)));
                continue;
            }

            if (IsIdentifierStart(current))
            {
                tokens.Add(ReadIdentifier(formula, ref index));
                continue;
            }

            switch (current)
            {
                case '(':
                    tokens.Add(new FormulaToken(FormulaTokenKind.LeftParenthesis, "("));
                    index++;
                    continue;
                case ')':
                    tokens.Add(new FormulaToken(FormulaTokenKind.RightParenthesis, ")"));
                    index++;
                    continue;
                case ',':
                    tokens.Add(new FormulaToken(FormulaTokenKind.Comma, ","));
                    index++;
                    continue;
                case ':':
                    tokens.Add(new FormulaToken(FormulaTokenKind.Colon, ":"));
                    index++;
                    continue;
                case '!':
                    tokens.Add(new FormulaToken(FormulaTokenKind.Bang, "!"));
                    index++;
                    continue;
                case '<':
                case '>':
                case '=':
                {
                    if (index + 1 < formula.Length && formula[index + 1] == '=')
                    {
                        tokens.Add(new FormulaToken(FormulaTokenKind.Operator, formula.Substring(index, 2)));
                        index += 2;
                        continue;
                    }

                    if (index + 1 < formula.Length && formula[index + 1] == '>' && current == '<')
                    {
                        tokens.Add(new FormulaToken(FormulaTokenKind.Operator, "<>"));
                        index += 2;
                        continue;
                    }

                    tokens.Add(new FormulaToken(FormulaTokenKind.Operator, current.ToString()));
                    index++;
                    continue;
                }

                case '+':
                case '-':
                case '*':
                case '/':
                case '^':
                case '&':
                case '%':
                    tokens.Add(new FormulaToken(FormulaTokenKind.Operator, current.ToString()));
                    index++;
                    continue;
            }

            throw new FormulaSyntaxException($"Unexpected character '{current}' at position {index}.");
        }

        tokens.Add(new FormulaToken(FormulaTokenKind.End, string.Empty));
        return tokens;
    }

    private static FormulaToken ReadNumber(string formula, ref int index)
    {
        int start = index;
        while (index < formula.Length && (char.IsAsciiDigit(formula[index]) || formula[index] == '.'))
        {
            index++;
        }

        if (index < formula.Length && (formula[index] == 'e' || formula[index] == 'E'))
        {
            int exponent = index + 1;
            if (exponent < formula.Length && (formula[exponent] == '+' || formula[exponent] == '-'))
            {
                exponent++;
            }

            if (exponent < formula.Length && char.IsAsciiDigit(formula[exponent]))
            {
                index = exponent;
                while (index < formula.Length && char.IsAsciiDigit(formula[index]))
                {
                    index++;
                }
            }
        }

        string text = formula[start..index];
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            throw new FormulaSyntaxException($"'{text}' is not a number.");
        }

        return new FormulaToken(FormulaTokenKind.Number, text, number);
    }

    private static FormulaToken ReadText(string formula, ref int index)
    {
        int start = ++index;
        StringBuilder builder = new();
        while (index < formula.Length)
        {
            if (formula[index] == '"')
            {
                if (index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    builder.Append('"');
                    index += 2;
                    continue;
                }

                index++;
                return new FormulaToken(FormulaTokenKind.Text, builder.ToString());
            }

            builder.Append(formula[index]);
            index++;
        }

        throw new FormulaSyntaxException($"The text that starts at position {start} is not closed.");
    }

    private static string ReadQuoted(string formula, ref int index)
    {
        int start = ++index;
        StringBuilder builder = new();
        while (index < formula.Length)
        {
            if (formula[index] == '\'')
            {
                if (index + 1 < formula.Length && formula[index + 1] == '\'')
                {
                    builder.Append('\'');
                    index += 2;
                    continue;
                }

                index++;
                return builder.ToString();
            }

            builder.Append(formula[index]);
            index++;
        }

        throw new FormulaSyntaxException($"The quoted name that starts at position {start} is not closed.");
    }

    private static FormulaToken ReadError(string formula, ref int index)
    {
        int start = index++;
        while (index < formula.Length && (char.IsAsciiLetterOrDigit(formula[index]) || formula[index] == '/' || formula[index] == '!'))
        {
            index++;
        }

        string code = formula[start..index];
        FormulaError? error = FormulaError.FromCode(code);
        if (error is null)
        {
            throw new FormulaSyntaxException($"'{code}' is not a known error value.");
        }

        return new FormulaToken(FormulaTokenKind.Error, code, error: error);
    }

    private static FormulaToken ReadIdentifier(string formula, ref int index)
    {
        int start = index;
        while (index < formula.Length && IsIdentifierPart(formula[index]))
        {
            index++;
        }

        string text = formula[start..index];
        if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase))
        {
            return new FormulaToken(FormulaTokenKind.Boolean, text, boolean: true);
        }

        if (string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return new FormulaToken(FormulaTokenKind.Boolean, text, boolean: false);
        }

        return new FormulaToken(FormulaTokenKind.Identifier, text);
    }

    internal static bool IsIdentifierStart(char value)
        => value is '_' or '$'
            || char.IsAsciiLetter(value)
            || value >= 0x80;

    internal static bool IsIdentifierPart(char value)
        => IsIdentifierStart(value) || char.IsAsciiDigit(value) || value == '.';
}
