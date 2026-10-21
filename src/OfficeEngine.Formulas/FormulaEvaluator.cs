using System.Globalization;

namespace OfficeEngine.Formulas;

/// <summary>Reads a formula and evaluates it over the values of a context.</summary>
internal sealed class FormulaEvaluator
{
    private readonly List<FormulaToken> _tokens;
    private readonly IFormulaContext _context;
    private readonly int _end;
    private int _position;

    internal FormulaEvaluator(List<FormulaToken> tokens, IFormulaContext context, int start = 0, int end = -1)
    {
        _tokens = tokens;
        _context = context;
        _position = start;
        _end = end < 0 ? tokens.Count - 1 : end;
    }

    internal object? Evaluate()
    {
        try
        {
            object? value = ParseComparison();
            if (_position < _end)
            {
                throw new FormulaSyntaxException($"Unexpected '{_tokens[_position].Text}' in the formula.");
            }

            return value;
        }
        catch (FormulaErrorException error)
        {
            return error.Error;
        }
    }

    private object? ParseComparison()
    {
        object? left = ParseConcatenation();
        while (Current.Kind == FormulaTokenKind.Operator && IsComparisonOperator(Current.Text))
        {
            string operation = Current.Text;
            Advance();
            object? right = ParseConcatenation();
            object? comparison = FormulaValues.CompareOrError(left, right);
            if (comparison is FormulaError)
            {
                left = comparison;
                continue;
            }

            int compared = (int)comparison!;
            bool result = operation switch
            {
                "=" => compared == 0,
                "<>" => compared != 0,
                "<" => compared < 0,
                ">" => compared > 0,
                "<=" => compared <= 0,
                _ => compared >= 0,
            };

            left = result;
        }

        return left;
    }

    private object? ParseConcatenation()
    {
        object? left = ParseAdditive();
        while (Current.IsOperator("&"))
        {
            Advance();
            object? right = ParseAdditive();
            object? first = FormulaValues.TextOrError(left);
            object? second = FormulaValues.TextOrError(right);
            if (first is FormulaError || second is FormulaError)
            {
                left = first is FormulaError ? first : second;
                continue;
            }

            left = (string)first! + (string)second!;
        }

        return left;
    }

    private object? ParseAdditive()
    {
        object? left = ParseMultiplicative();
        while (Current.IsOperator("+") || Current.IsOperator("-"))
        {
            bool add = Current.Text == "+";
            Advance();
            object? right = ParseMultiplicative();
            object? firstValue = FormulaValues.NumberOrError(left);
            object? secondValue = FormulaValues.NumberOrError(right);
            if (firstValue is FormulaError || secondValue is FormulaError)
            {
                left = firstValue is FormulaError ? firstValue : secondValue;
                continue;
            }

            double first = (double)firstValue!;
            double second = (double)secondValue!;
            left = add ? first + second : first - second;
        }

        return left;
    }

    private object? ParseMultiplicative()
    {
        object? left = ParseUnary();
        while (Current.IsOperator("*") || Current.IsOperator("/"))
        {
            bool multiply = Current.Text == "*";
            Advance();
            object? right = ParseUnary();
            object? firstValue = FormulaValues.NumberOrError(left);
            object? secondValue = FormulaValues.NumberOrError(right);
            if (firstValue is FormulaError || secondValue is FormulaError)
            {
                left = firstValue is FormulaError ? firstValue : secondValue;
                continue;
            }

            double first = (double)firstValue!;
            double second = (double)secondValue!;
            if (!multiply && second == 0d)
            {
                left = FormulaError.DivideByZero;
                continue;
            }

            left = multiply ? first * second : first / second;
        }

        return left;
    }

    private object? ParseUnary()
    {
        if (Current.IsOperator("-"))
        {
            Advance();
            object? value = FormulaValues.NumberOrError(ParseUnary());
            return value is FormulaError ? value : -(double)value!;
        }

        if (Current.IsOperator("+"))
        {
            Advance();
            return FormulaValues.NumberOrError(ParseUnary());
        }

        return ParsePower();
    }

    private object? ParsePower()
    {
        object? left = ParsePostfix();
        if (Current.IsOperator("^"))
        {
            Advance();
            object? right = ParseUnary();
            object? firstValue = FormulaValues.NumberOrError(left);
            object? secondValue = FormulaValues.NumberOrError(right);
            if (firstValue is FormulaError || secondValue is FormulaError)
            {
                left = firstValue is FormulaError ? firstValue : secondValue;
                return left;
            }

            left = Math.Pow((double)firstValue!, (double)secondValue!);
        }

        return left;
    }

    private object? ParsePostfix()
    {
        object? value = ParsePrimary();
        while (Current.IsOperator("%"))
        {
            Advance();
            object? number = FormulaValues.NumberOrError(value);
            if (number is FormulaError)
            {
                value = number;
                continue;
            }

            value = (double)number! / 100d;
        }

        return value;
    }

    private object? ParsePrimary()
    {
        FormulaToken token = Current;
        switch (token.Kind)
        {
            case FormulaTokenKind.Number:
                Advance();
                return token.Number;
            case FormulaTokenKind.Text:
                Advance();
                return token.Text;
            case FormulaTokenKind.Boolean:
                Advance();
                return token.Boolean;
            case FormulaTokenKind.Error:
                Advance();
                return token.Error;
            case FormulaTokenKind.LeftParenthesis:
            {
                Advance();
                object? value = ParseComparison();
                Expect(FormulaTokenKind.RightParenthesis, ")");
                return value;
            }

            case FormulaTokenKind.Identifier:
                return ParseIdentifier();
        }

        throw new FormulaSyntaxException($"Unexpected '{token.Text}' in the formula.");
    }

    private object? ParseIdentifier()
    {
        FormulaToken token = Current;

        // A sheet qualified reference: Sheet1!A1 or 'My Sheet'!A1.
        if (Peek(1).Kind == FormulaTokenKind.Bang)
        {
            string sheetName = token.Text;
            Advance();
            Advance();
            return ParseCellOrRange(sheetName);
        }

        // A range: A1:B2.
        if (Peek(1).Kind == FormulaTokenKind.Colon && TryParseCellReference(token.Text, out _, out _))
        {
            return ParseCellOrRange(null);
        }

        // A function call.
        if (Peek(1).Kind == FormulaTokenKind.LeftParenthesis)
        {
            return ParseFunctionCall(token.Text);
        }

        // A plain cell reference.
        if (TryParseCellReference(token.Text, out _, out _))
        {
            Advance();
            return ReadCell(null, token.Text);
        }

        // An unknown name.
        Advance();
        throw new FormulaErrorException(FormulaError.UnknownName);
    }

    private object? ParseCellOrRange(string? sheetName)
    {
        FormulaToken token = Current;
        if (token.Kind != FormulaTokenKind.Identifier || !TryParseCellReference(token.Text, out int row, out int column))
        {
            throw new FormulaSyntaxException($"'{token.Text}' is not a cell reference.");
        }

        Advance();

        if (Current.Kind == FormulaTokenKind.Colon)
        {
            Advance();
            FormulaToken second = Current;
            if (second.Kind != FormulaTokenKind.Identifier || !TryParseCellReference(second.Text, out int lastRow, out int lastColumn))
            {
                throw new FormulaSyntaxException($"'{second.Text}' is not a cell reference.");
            }

            Advance();
            return new FormulaRange(
                sheetName,
                Math.Min(row, lastRow),
                Math.Min(column, lastColumn),
                Math.Max(row, lastRow),
                Math.Max(column, lastColumn));
        }

        return ReadCell(sheetName, token.Text, row, column);
    }

    private object? ParseFunctionCall(string name)
    {
        Advance();
        Advance();

        // The Spanish name of a function is an alias of the canonical one (ADR 0043).
        string canonical = FormulaNames.CanonicalName(name);

        if (string.Equals(canonical, "IF", StringComparison.OrdinalIgnoreCase))
        {
            return ParseIf();
        }

        if (string.Equals(canonical, "IFERROR", StringComparison.OrdinalIgnoreCase))
        {
            return ParseIfError();
        }

        List<object?> arguments = [];
        if (Current.Kind != FormulaTokenKind.RightParenthesis)
        {
            while (true)
            {
                arguments.Add(ParseComparison());
                if (Current.Kind == FormulaTokenKind.Comma)
                {
                    Advance();
                    continue;
                }

                break;
            }
        }

        Expect(FormulaTokenKind.RightParenthesis, ")");
        return FormulaFunctions.Invoke(canonical, arguments.ToArray(), _context);
    }

    /// <summary>Reads <c>IF</c> without evaluating the branch that is not taken.</summary>
    private object? ParseIf()
    {
        int conditionStart = _position;
        int conditionEnd = FindArgumentEnd(conditionStart);
        int thenStart = conditionEnd + 1;
        int thenEnd = FindArgumentEnd(thenStart);
        int elseStart = thenEnd + 1;
        int elseEnd = FindArgumentEnd(elseStart);

        bool taken = FormulaValues.ToBoolean(EvaluateSub(conditionStart, conditionEnd));
        object? result = taken
            ? EvaluateSub(thenStart, thenEnd)
            : elseStart < elseEnd
                ? EvaluateSub(elseStart, elseEnd)
                : false;

        _position = elseEnd + 1;
        return result;
    }

    /// <summary>Reads <c>IFERROR</c> without failing when the first argument is an error.</summary>
    private object? ParseIfError()
    {
        int firstStart = _position;
        int firstEnd = FindArgumentEnd(firstStart);
        int secondStart = firstEnd + 1;
        int secondEnd = FindArgumentEnd(secondStart);

        object? value = EvaluateSub(firstStart, firstEnd);
        if (value is FormulaError)
        {
            value = EvaluateSub(secondStart, secondEnd);
        }

        _position = secondEnd + 1;
        return value;
    }

    private object? EvaluateSub(int start, int end)
    {
        if (start >= end)
        {
            return null;
        }

        FormulaEvaluator sub = new(_tokens, _context, start, end);
        return sub.Evaluate();
    }

    /// <summary>Finds where an argument ends: the comma that follows it, or its closing parenthesis.</summary>
    private int FindArgumentEnd(int start)
    {
        int depth = 0;
        int index = start;
        while (index < _tokens.Count)
        {
            FormulaToken token = _tokens[index];
            if (token.Kind == FormulaTokenKind.LeftParenthesis)
            {
                depth++;
            }
            else if (token.Kind == FormulaTokenKind.RightParenthesis)
            {
                if (depth == 0)
                {
                    return index;
                }

                depth--;
            }
            else if (token.Kind == FormulaTokenKind.Comma && depth == 0)
            {
                return index;
            }
            else if (token.Kind == FormulaTokenKind.End)
            {
                return index;
            }

            index++;
        }

        return index;
    }

    private object? ReadCell(string? sheetName, string reference)
        => ReadCell(sheetName, reference, 0, 0);

    private object? ReadCell(string? sheetName, string reference, int row, int column)
    {
        if (row == 0 || column == 0)
        {
            if (!TryParseCellReference(reference, out row, out column))
            {
                throw new FormulaSyntaxException($"'{reference}' is not a cell reference.");
            }
        }

        _context.TryGetCellValue(sheetName ?? _context.SheetName, row, column, out object? value);
        return value;
    }

    private static bool TryParseCellReference(string text, out int row, out int column)
    {
        row = 0;
        column = 0;
        ReadOnlySpan<char> span = text.AsSpan();
        int index = 0;
        if (index < span.Length && span[index] == '$')
        {
            index++;
        }

        int letterStart = index;
        while (index < span.Length && char.IsAsciiLetter(span[index]))
        {
            index++;
        }

        if (index == letterStart)
        {
            return false;
        }

        int letterEnd = index;
        if (index < span.Length && span[index] == '$')
        {
            index++;
        }

        int digitStart = index;
        while (index < span.Length && char.IsAsciiDigit(span[index]))
        {
            index++;
        }

        if (index == digitStart || index != span.Length)
        {
            return false;
        }

        long columnValue = 0;
        for (int letter = letterStart; letter < letterEnd; letter++)
        {
            columnValue = (columnValue * 26) + (char.ToUpperInvariant(span[letter]) - 'A' + 1);
            if (columnValue > 16384)
            {
                return false;
            }
        }

        if (!long.TryParse(span[digitStart..index], NumberStyles.None, CultureInfo.InvariantCulture, out long rowValue)
            || rowValue < 1
            || rowValue > 1048576)
        {
            return false;
        }

        column = (int)columnValue;
        row = (int)rowValue;
        return true;
    }

    private static bool IsComparisonOperator(string value)
        => value is "=" or "<>" or "<" or ">" or "<=" or ">=";

    private FormulaToken Current => _tokens[_position];

    private FormulaToken Peek(int offset)
    {
        int index = _position + offset;
        return index < _tokens.Count ? _tokens[index] : _tokens[^1];
    }

    private void Advance() => _position++;

    private void Expect(FormulaTokenKind kind, string text)
    {
        if (Current.Kind != kind)
        {
            throw new FormulaSyntaxException($"Expected '{text}' in the formula.");
        }

        Advance();
    }
}
