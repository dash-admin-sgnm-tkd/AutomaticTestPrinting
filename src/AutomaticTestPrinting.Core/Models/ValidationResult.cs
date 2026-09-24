namespace AutomaticTestPrinting.Core.Models;

public sealed record ValidationResult(bool IsValid, string Message)
{
    public static ValidationResult Success() => new(true, string.Empty);

    public static ValidationResult Failure(string message) => new(false, message);
}
