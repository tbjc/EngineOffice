using System.Text;
using System.Globalization;

namespace OfficeEngine.Formulas;

/// <summary>One function of the matrix, with the range of arguments it accepts.</summary>
internal sealed class FormulaFunction
{
    internal FormulaFunction(
        string name,
        int minimumArguments,
        int maximumArguments,
        Func<object?[], IFormulaContext, object?> implementation,
        bool acceptsErrors = false)
    {
        Name = name;
        MinimumArguments = minimumArguments;
        MaximumArguments = maximumArguments;
        Implementation = implementation;
        AcceptsErrors = acceptsErrors;
    }

    internal string Name { get; }

    internal int MinimumArguments { get; }

    internal int MaximumArguments { get; }

    internal Func<object?[], IFormulaContext, object?> Implementation { get; }

    /// <summary>Gets whether the function receives the errors of its arguments instead of failing.</summary>
    internal bool AcceptsErrors { get; }
}

/// <summary>The functions this engine understands.</summary>
internal static class FormulaFunctions
{
    private static readonly Dictionary<string, FormulaFunction> Functions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["SUM"] = new("SUM", 1, int.MaxValue, (args, context) => Numbers(args, context).Sum()),
            ["PRODUCT"] = new("PRODUCT", 1, int.MaxValue, (args, context) =>
            {
                double product = 1d;
                foreach (double number in Numbers(args, context))
                {
                    product *= number;
                }

                return product;
            }),
            ["AVERAGE"] = new("AVERAGE", 1, int.MaxValue, (args, context) =>
            {
                List<double> numbers = Numbers(args, context);
                if (numbers.Count == 0)
                {
                    throw new FormulaErrorException(FormulaError.DivideByZero);
                }

                return numbers.Sum() / numbers.Count;
            }),
            ["MIN"] = new("MIN", 1, int.MaxValue, (args, context) =>
            {
                List<double> numbers = Numbers(args, context);
                return numbers.Count == 0 ? 0d : numbers.Min();
            }),
            ["MAX"] = new("MAX", 1, int.MaxValue, (args, context) =>
            {
                List<double> numbers = Numbers(args, context);
                return numbers.Count == 0 ? 0d : numbers.Max();
            }),
            ["COUNT"] = new("COUNT", 1, int.MaxValue, (args, context) => (double)CountNumbers(args, context)),
            ["COUNTA"] = new("COUNTA", 1, int.MaxValue, (args, context) => (double)CountValues(args, context)),
            ["AND"] = new("AND", 1, int.MaxValue, (args, context) => Booleans(args, context).All(value => value)),
            ["OR"] = new("OR", 1, int.MaxValue, (args, context) => Booleans(args, context).Any(value => value)),
            ["NOT"] = new("NOT", 1, 1, (args, _) => !FormulaValues.ToBoolean(args[0])),
            ["ABS"] = new("ABS", 1, 1, (args, _) => Math.Abs(FormulaValues.ToNumber(args[0]))),
            ["INT"] = new("INT", 1, 1, (args, _) => Math.Truncate(FormulaValues.ToNumber(args[0]))),
            ["SQRT"] = new("SQRT", 1, 1, (args, _) =>
            {
                double value = FormulaValues.ToNumber(args[0]);
                if (value < 0d)
                {
                    throw new FormulaErrorException(FormulaError.Number);
                }

                return Math.Sqrt(value);
            }),
            ["ROUND"] = new("ROUND", 1, 2, (args, _) => Round(FormulaValues.ToNumber(args[0]), Digits(args, 1), MidpointRounding.AwayFromZero)),
            ["ROUNDUP"] = new("ROUNDUP", 1, 2, (args, _) => RoundUp(FormulaValues.ToNumber(args[0]), Digits(args, 1))),
            ["ROUNDDOWN"] = new("ROUNDDOWN", 1, 2, (args, _) => RoundDown(FormulaValues.ToNumber(args[0]), Digits(args, 1))),
            ["MOD"] = new("MOD", 2, 2, (args, _) =>
            {
                double number = FormulaValues.ToNumber(args[0]);
                double divisor = FormulaValues.ToNumber(args[1]);
                if (divisor == 0d)
                {
                    throw new FormulaErrorException(FormulaError.DivideByZero);
                }

                return number - (divisor * Math.Floor(number / divisor));
            }),
            ["POWER"] = new("POWER", 2, 2, (args, _) => Math.Pow(FormulaValues.ToNumber(args[0]), FormulaValues.ToNumber(args[1]))),
            ["CONCATENATE"] = new("CONCATENATE", 1, int.MaxValue, (args, _) =>
                string.Concat(args.Select(FormulaValues.ToText))),
            ["LEN"] = new("LEN", 1, 1, (args, _) => (double)FormulaValues.ToText(args[0]).Length),
            ["LEFT"] = new("LEFT", 1, 2, (args, _) => Left(FormulaValues.ToText(args[0]), Digits(args, 1))),
            ["RIGHT"] = new("RIGHT", 1, 2, (args, _) => Right(FormulaValues.ToText(args[0]), Digits(args, 1))),
            ["MID"] = new("MID", 3, 3, (args, _) =>
            {
                string text = FormulaValues.ToText(args[0]);
                int start = (int)FormulaValues.ToNumber(args[1]);
                int length = (int)FormulaValues.ToNumber(args[2]);
                if (start < 1 || length < 0)
                {
                    throw new FormulaErrorException(FormulaError.Value);
                }

                if (start > text.Length)
                {
                    return string.Empty;
                }

                return text.Substring(start - 1, Math.Min(length, text.Length - start + 1));
            }),
            ["UPPER"] = new("UPPER", 1, 1, (args, _) => FormulaValues.ToText(args[0]).ToUpperInvariant()),
            ["LOWER"] = new("LOWER", 1, 1, (args, _) => FormulaValues.ToText(args[0]).ToLowerInvariant()),
            ["TRIM"] = new("TRIM", 1, 1, (args, _) => Trim(FormulaValues.ToText(args[0]))),
            ["ISBLANK"] = new("ISBLANK", 1, 1, (args, _) => args[0] is null),
            ["ISNUMBER"] = new("ISNUMBER", 1, 1, (args, _) => args[0] is double),
            ["ISTEXT"] = new("ISTEXT", 1, 1, (args, _) => args[0] is string),
            ["ISERROR"] = new("ISERROR", 1, 1, (args, _) => args[0] is FormulaError, acceptsErrors: true),
        };

    /// <summary>Gets the names of the functions this engine understands, in a stable order.</summary>
    internal static IReadOnlyList<string> Names { get; } =
        [.. Functions.Keys
            .Concat(["IF", "IFERROR"])
            .Concat(FormulaNames.SpanishAliases)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.Ordinal)];

    /// <summary>Runs a function with its arguments already evaluated.</summary>
    internal static object? Invoke(string name, object?[] arguments, IFormulaContext context)
    {
        if (!Functions.TryGetValue(name, out FormulaFunction? function))
        {
            throw new FormulaErrorException(FormulaError.UnknownName);
        }

        if (arguments.Length < function.MinimumArguments
            || (function.MaximumArguments != int.MaxValue && arguments.Length > function.MaximumArguments))
        {
            throw new FormulaErrorException(FormulaError.Value);
        }

        if (!function.AcceptsErrors)
        {
            foreach (object? argument in arguments)
            {
                if (argument is FormulaError error)
                {
                    return error;
                }
            }
        }

        return function.Implementation(arguments, context);
    }

    private static List<double> Numbers(object?[] arguments, IFormulaContext context)
    {
        List<double> numbers = [];
        foreach (object? argument in arguments)
        {
            if (argument is FormulaRange range)
            {
                foreach (object? value in FormulaRanges.Values(range, context))
                {
                    if (value is double number)
                    {
                        numbers.Add(number);
                    }
                }

                continue;
            }

            if (argument is bool flag)
            {
                numbers.Add(flag ? 1d : 0d);
                continue;
            }

            if (argument is string text)
            {
                if (text.Trim().Length == 0)
                {
                    continue;
                }

                numbers.Add(FormulaValues.ToNumber(text));
                continue;
            }

            if (argument is null)
            {
                continue;
            }

            numbers.Add(FormulaValues.ToNumber(argument));
        }

        return numbers;
    }

    private static List<bool> Booleans(object?[] arguments, IFormulaContext context)
    {
        List<bool> booleans = [];
        foreach (object? argument in arguments)
        {
            if (argument is FormulaRange range)
            {
                foreach (object? value in FormulaRanges.Values(range, context))
                {
                    if (value is bool flag)
                    {
                        booleans.Add(flag);
                    }
                    else if (value is double number)
                    {
                        booleans.Add(number != 0d);
                    }
                }

                continue;
            }

            booleans.Add(FormulaValues.ToBoolean(argument));
        }

        return booleans;
    }

    private static int CountNumbers(object?[] arguments, IFormulaContext context)
    {
        int count = 0;
        foreach (object? argument in arguments)
        {
            if (argument is FormulaRange range)
            {
                foreach (object? value in FormulaRanges.Values(range, context))
                {
                    if (value is double)
                    {
                        count++;
                    }
                }

                continue;
            }

            if (argument is bool or double)
            {
                count++;
                continue;
            }

            if (argument is string text
                && double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out _))
            {
                count++;
            }
        }

        return count;
    }

    private static int CountValues(object?[] arguments, IFormulaContext context)
    {
        int count = 0;
        foreach (object? argument in arguments)
        {
            if (argument is FormulaRange range)
            {
                foreach (object? value in FormulaRanges.Values(range, context))
                {
                    if (value is not null)
                    {
                        count++;
                    }
                }

                continue;
            }

            if (argument is not null)
            {
                count++;
            }
        }

        return count;
    }

    private static int Digits(object?[] arguments, int index)
        => index < arguments.Length ? (int)FormulaValues.ToNumber(arguments[index]) : 0;

    private static double Round(double value, int digits, MidpointRounding mode)
        => Math.Round(value, Math.Clamp(digits, -15, 15), mode);

    private static double RoundUp(double value, int digits)
    {
        double factor = Math.Pow(10d, Math.Clamp(digits, -15, 15));
        double scaled = value * factor;
        double result = scaled >= 0d ? Math.Ceiling(scaled) : Math.Floor(scaled);
        return result / factor;
    }

    private static double RoundDown(double value, int digits)
    {
        double factor = Math.Pow(10d, Math.Clamp(digits, -15, 15));
        double scaled = value * factor;
        double result = scaled >= 0d ? Math.Floor(scaled) : Math.Ceiling(scaled);
        return result / factor;
    }

    private static string Left(string text, int count)
    {
        if (count <= 0)
        {
            return string.Empty;
        }

        return count >= text.Length ? text : text[..count];
    }

    private static string Right(string text, int count)
    {
        if (count <= 0)
        {
            return string.Empty;
        }

        return count >= text.Length ? text : text[^count..];
    }

    private static string Trim(string text)
    {
        ReadOnlySpan<char> span = text.AsSpan().Trim();
        bool previousWasSpace = false;
        StringBuilder builder = new();
        foreach (char character in span)
        {
            bool isSpace = char.IsWhiteSpace(character);
            if (isSpace && previousWasSpace)
            {
                continue;
            }

            builder.Append(character);
            previousWasSpace = isSpace;
        }

        return builder.ToString();
    }
}
