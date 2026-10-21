namespace OfficeEngine.Diagnostics;

/// <summary>
/// Indicates the importance of an <see cref="OfficeEngineDiagnostic"/>.
/// </summary>
public enum DiagnosticSeverity
{
    /// <summary>Informational message, such as content that is preserved but not modelled.</summary>
    Info,

    /// <summary>Something the caller should be aware of, such as a possible loss of fidelity.</summary>
    Warning,

    /// <summary>A problem that prevented an operation from completing correctly.</summary>
    Error,
}

/// <summary>
/// A single message produced while loading or processing a document.
/// </summary>
/// <remarks>
/// Diagnostics are the mechanism OfficeEngine uses to avoid silent data loss: content
/// that cannot be modelled is reported here rather than dropped without notice.
/// </remarks>
public sealed class OfficeEngineDiagnostic
{
    /// <summary>Initializes a new instance of the <see cref="OfficeEngineDiagnostic"/> class.</summary>
    /// <param name="severity">The severity of the message.</param>
    /// <param name="code">A stable, machine-readable code for the message.</param>
    /// <param name="message">A human-readable description.</param>
    public OfficeEngineDiagnostic(DiagnosticSeverity severity, string code, string message)
    {
        ArgumentException.ThrowIfNullOrEmpty(code);
        ArgumentException.ThrowIfNullOrEmpty(message);

        Severity = severity;
        Code = code;
        Message = message;
    }

    /// <summary>Gets the severity of the message.</summary>
    public DiagnosticSeverity Severity { get; }

    /// <summary>Gets the stable, machine-readable code for the message.</summary>
    public string Code { get; }

    /// <summary>Gets the human-readable description.</summary>
    public string Message { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Severity} {Code}: {Message}";
}

/// <summary>
/// An ordered accumulator of <see cref="OfficeEngineDiagnostic"/> values.
/// </summary>
/// <remarks>
/// Consumers receive diagnostics through <c>IReadOnlyList&lt;OfficeEngineDiagnostic&gt;</c>
/// properties on documents. This type exists so the library can accumulate messages while loading;
/// use <see cref="ToSnapshot"/> to obtain an immutable view.
/// </remarks>
public sealed class OfficeEngineDiagnostics
{
    private readonly List<OfficeEngineDiagnostic> _items = [];

    /// <summary>Gets the number of diagnostics collected.</summary>
    public int Count => _items.Count;

    /// <summary>Gets a value indicating whether any diagnostic has <see cref="DiagnosticSeverity.Warning"/> severity.</summary>
    public bool HasWarnings
    {
        get
        {
            foreach (OfficeEngineDiagnostic item in _items)
            {
                if (item.Severity == DiagnosticSeverity.Warning)
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>Gets a value indicating whether any diagnostic has <see cref="DiagnosticSeverity.Error"/> severity.</summary>
    public bool HasErrors
    {
        get
        {
            foreach (OfficeEngineDiagnostic item in _items)
            {
                if (item.Severity == DiagnosticSeverity.Error)
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>Adds a diagnostic message.</summary>
    /// <param name="diagnostic">The message to add.</param>
    /// <exception cref="ArgumentNullException"><paramref name="diagnostic"/> is <see langword="null"/>.</exception>
    public void Add(OfficeEngineDiagnostic diagnostic)
    {
        ArgumentNullException.ThrowIfNull(diagnostic);
        _items.Add(diagnostic);
    }

    /// <summary>Adds a diagnostic message built from its parts.</summary>
    /// <param name="severity">The severity of the message.</param>
    /// <param name="code">A stable, machine-readable code for the message.</param>
    /// <param name="message">A human-readable description.</param>
    public void Add(DiagnosticSeverity severity, string code, string message)
        => Add(new OfficeEngineDiagnostic(severity, code, message));

    /// <summary>Removes every collected diagnostic.</summary>
    public void Clear() => _items.Clear();

    /// <summary>Gets an immutable snapshot of the collected diagnostics.</summary>
    /// <returns>A snapshot that is not affected by later additions.</returns>
    public IReadOnlyList<OfficeEngineDiagnostic> ToSnapshot() => _items.ToArray();
}
