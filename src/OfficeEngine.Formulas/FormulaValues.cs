using System.Globalization;

namespace OfficeEngine.Formulas;

/// <summary>Thrown while a formula is evaluated, to carry an error value.</summary>
internal sealed class FormulaErrorException : Exception
{
    internal FormulaErrorException(FormulaError error)
        : base(error.Code)
        => Error = error;

    internal FormulaError Error { get; }
}

/// <summary>The conversions between the value types of a formula, with the rules of Excel.</summary>
internal static class FormulaValues
{
    /// <summary>Gets whether the value is an error.</summary>
    internal static bool IsError(object? value) => value is FormulaError;

    /// <summary>Throws when the value is an error, so that the error travels up the expression.</summary>
    internal static object? Propagate(object? value)
    {
        if (value is FormulaError error)
        {
            throw new FormulaErrorException(error);
        }

        return value;
    }

    /// <summary>Converts a value to a number, or fails with <c>#VALUE!</c>.</summary>
    internal static double ToNumber(object? value)
    {
        object? propagated = Propagate(value);
        return propagated switch
        {
            null => 0d,
            double number => number,
            bool flag => flag ? 1d : 0d,
            string text => ParseNumber(text),
            FormulaRange => throw new FormulaErrorException(FormulaError.Value),
            _ => throw new FormulaErrorException(FormulaError.Value),
        };
    }

    /// <summary>
    /// Converts a value to a number, keeping the error instead of raising it, so that the caller can
    /// decide what to do with it.
    /// </summary>
    internal static object? NumberOrError(object? value)
    {
        try
        {
            return ToNumber(value);
        }
        catch (FormulaErrorException error)
        {
            return error.Error;
        }
    }

    /// <summary>Converts a value to text, keeping the error instead of raising it.</summary>
    internal static object? TextOrError(object? value)
    {
        try
        {
            return ToText(value);
        }
        catch (FormulaErrorException error)
        {
            return error.Error;
        }
    }

    /// <summary>Converts a value to a boolean, keeping the error instead of raising it.</summary>
    internal static object? BooleanOrError(object? value)
    {
        try
        {
            return ToBoolean(value);
        }
        catch (FormulaErrorException error)
        {
            return error.Error;
        }
    }

    /// <summary>Converts a value to text, or fails with <c>#VALUE!</c>.</summary>
    internal static string ToText(object? value)
    {
        object? propagated = Propagate(value);
        return propagated switch
        {
            null => string.Empty,
            double number => NumberToText(number),
            bool flag => flag ? "TRUE" : "FALSE",
            string text => text,
            FormulaRange => throw new FormulaErrorException(FormulaError.Value),
            _ => throw new FormulaErrorException(FormulaError.Value),
        };
    }

    /// <summary>Converts a value to a boolean, or fails with <c>#VALUE!</c>.</summary>
    internal static bool ToBoolean(object? value)
    {
        object? propagated = Propagate(value);
        return propagated switch
        {
            null => false,
            bool flag => flag,
            double number => number != 0d,
            string text => ParseBoolean(text),
            FormulaRange => throw new FormulaErrorException(FormulaError.Value),
            _ => throw new FormulaErrorException(FormulaError.Value),
        };
    }

    /// <summary>Formats a number the way the general format of Excel does.</summary>
    internal static string NumberToText(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            throw new FormulaErrorException(FormulaError.Number);
        }

        return value.ToString("G15", CultureInfo.InvariantCulture);
    }

    /// <summary>Compares two values, keeping the error instead of raising it.</summary>
    internal static object? CompareOrError(object? left, object? right)
    {
        try
        {
            return Compare(left, right);
        }
        catch (FormulaErrorException error)
        {
            return error.Error;
        }
    }

    /// <summary>Compares two values with the order of Excel: numbers, then text, then booleans.</summary>
    internal static int Compare(object? left, object? right)
    {
        object? first = Propagate(left);
        object? second = Propagate(right);
        if (first is FormulaRange || second is FormulaRange)
        {
            throw new FormulaErrorException(FormulaError.Value);
        }

        int rankLeft = Rank(first);
        int rankRight = Rank(second);
        if (rankLeft != rankRight)
        {
            return rankLeft < rankRight ? -1 : 1;
        }

        return rankLeft switch
        {
            0 => first is double a && second is double b
                ? a.CompareTo(b)
                : 0,
            1 => string.Compare(
                (string?)first ?? string.Empty,
                (string?)second ?? string.Empty,
                StringComparison.OrdinalIgnoreCase),
            _ => ((bool?)first ?? false).CompareTo((bool?)second ?? false),
        };
    }

    private static int Rank(object? value) => value switch
    {
        null => 0,
        double => 0,
        string => 1,
        bool => 2,
        _ => 0,
    };

    private static double ParseNumber(string text)
    {
        string trimmed = text.Trim();
        if (trimmed.Length == 0)
        {
            throw new FormulaErrorException(FormulaError.Value);
        }

        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            return number;
        }

        if (DateTime.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
        {
            return date.ToOADate();
        }

        throw new FormulaErrorException(FormulaError.Value);
    }

    private static bool ParseBoolean(string text)
    {
        if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        throw new FormulaErrorException(FormulaError.Value);
    }
}
