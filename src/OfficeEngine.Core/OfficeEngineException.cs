using System.Diagnostics.CodeAnalysis;

namespace OfficeEngine;

/// <summary>
/// Base class for every exception raised by OfficeEngine.
/// </summary>
/// <remarks>
/// Callers can catch this type to handle all library-specific failures while still
/// being able to distinguish them from general BCL exceptions.
/// </remarks>
[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "The serialization constructor is obsolete since .NET 8 (SYSLIB0051) and OfficeEngine exceptions are not serialized.")]
public abstract class OfficeEngineException : Exception
{
    /// <summary>Initializes a new instance of the <see cref="OfficeEngineException"/> class.</summary>
    protected OfficeEngineException()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="OfficeEngineException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    protected OfficeEngineException(string message)
        : base(message)
    {
    }

    /// <summary>Initializes a new instance of the <see cref="OfficeEngineException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that caused the current exception.</param>
    protected OfficeEngineException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>
/// Thrown when an input is not a valid, readable Office Open XML file.
/// </summary>
public sealed class InvalidOfficeFileException : OfficeEngineException
{
    /// <summary>Initializes a new instance of the <see cref="InvalidOfficeFileException"/> class.</summary>
    public InvalidOfficeFileException()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="InvalidOfficeFileException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    public InvalidOfficeFileException(string message)
        : base(message)
    {
    }

    /// <summary>Initializes a new instance of the <see cref="InvalidOfficeFileException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that caused the current exception.</param>
    public InvalidOfficeFileException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>
/// Thrown when an operation is not supported by the current implementation.
/// </summary>
public sealed class UnsupportedOfficeFeatureException : OfficeEngineException
{
    /// <summary>Initializes a new instance of the <see cref="UnsupportedOfficeFeatureException"/> class.</summary>
    public UnsupportedOfficeFeatureException()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="UnsupportedOfficeFeatureException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    public UnsupportedOfficeFeatureException(string message)
        : base(message)
    {
    }

    /// <summary>Initializes a new instance of the <see cref="UnsupportedOfficeFeatureException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that caused the current exception.</param>
    public UnsupportedOfficeFeatureException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>
/// Thrown when reading an untrusted Office file would exceed a configured resource limit.
/// </summary>
public sealed class OfficeResourceLimitExceededException : OfficeEngineException
{
    /// <summary>Initializes a new instance of the <see cref="OfficeResourceLimitExceededException"/> class.</summary>
    public OfficeResourceLimitExceededException()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="OfficeResourceLimitExceededException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    public OfficeResourceLimitExceededException(string message)
        : base(message)
    {
    }

    /// <summary>Initializes a new instance of the <see cref="OfficeResourceLimitExceededException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">The exception that caused the current exception.</param>
    public OfficeResourceLimitExceededException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
