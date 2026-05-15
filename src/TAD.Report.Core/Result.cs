namespace TAD.Report.Core;

/// <summary>
/// Factory helpers for <see cref="Result{T}"/> used by infrastructure APIs (e.g. <c>Result&lt;byte[]&gt;</c>).
/// </summary>
public static class Result
{
    public static Result<T> Success<T>(T value) => Result<T>.Success(value);

    public static Result<T> Failure<T>(string errorMessage) => Result<T>.Failure(errorMessage);
}

/// <summary>
/// Unified operation outcome: success carries <see cref="Value"/>; failure carries <see cref="ErrorMessage"/>.
/// </summary>
/// <typeparam name="T">Payload type on success.</typeparam>
public readonly struct Result<T>
{
    private Result(bool isSuccess, string? errorMessage, T? value)
    {
        IsSuccess = isSuccess;
        ErrorMessage = errorMessage;
        Value = value;
    }

    /// <summary>True when the operation completed without a logical error.</summary>
    public bool IsSuccess { get; }

    /// <summary>Human-readable error description when <see cref="IsSuccess"/> is false.</summary>
    public string? ErrorMessage { get; }

    /// <summary>Result payload when <see cref="IsSuccess"/> is true.</summary>
    public T? Value { get; }

    public static Result<T> Success(T value) => new(true, null, value);

    public static Result<T> Failure(string errorMessage) =>
        new(false, errorMessage ?? string.Empty, default);
}
