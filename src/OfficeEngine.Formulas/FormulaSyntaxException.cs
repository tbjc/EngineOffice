namespace OfficeEngine.Formulas;

/// <summary>Thrown when a formula cannot be read.</summary>
public sealed class FormulaSyntaxException : Exception
{
    /// <summary>Initializes a new instance of the <see cref="FormulaSyntaxException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    public FormulaSyntaxException(string message)
        : base(message)
    {
    }
}
