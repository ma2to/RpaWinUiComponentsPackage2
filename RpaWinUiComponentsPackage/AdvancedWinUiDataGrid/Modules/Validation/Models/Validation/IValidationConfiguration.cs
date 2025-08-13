namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Validation.Models;

/// <summary>
/// Interface pre validation configuration - implementuje sa v aplikácii, NIE v balíku
/// Aplikácia definuje svoje vlastné business validation rules
/// </summary>
public interface IValidationConfiguration
{
    /// <summary>
    /// Získa validation rules pre všetky stĺpce
    /// Implementuje sa v aplikácii s vlastnými business rules
    /// </summary>
    ValidationRuleSet GetValidationRules();

    /// <summary>
    /// Získa cross-row validation rules (rules ktoré validujú celý dataset)
    /// </summary>
    List<CrossRowValidationRule> GetCrossRowValidationRules();

    /// <summary>
    /// Je validation zapnutá
    /// </summary>
    bool IsValidationEnabled { get; }

    /// <summary>
    /// Má sa spúšťať real-time validation (pri typing)
    /// </summary>
    bool EnableRealtimeValidation { get; }

    /// <summary>
    /// Má sa spúšťať batch validation (pri import/paste)
    /// </summary>
    bool EnableBatchValidation { get; }
}

/// <summary>
/// Set validation rules pre všetky stĺpce
/// </summary>
public class ValidationRuleSet
{
    private readonly Dictionary<string, List<ValidationRule>> _columnRules = new();

    /// <summary>
    /// Pridá validation rule pre stĺpec
    /// </summary>
    public void AddRule(string columnName, ValidationRule rule)
    {
        if (!_columnRules.ContainsKey(columnName))
            _columnRules[columnName] = new List<ValidationRule>();

        _columnRules[columnName].Add(rule);
    }

    /// <summary>
    /// Získa validation rules pre stĺpec
    /// </summary>
    public List<ValidationRule> GetRules(string columnName)
    {
        return _columnRules.TryGetValue(columnName, out var rules) ? rules : new List<ValidationRule>();
    }

    /// <summary>
    /// Získa všetky column names s validation rules
    /// </summary>
    public IEnumerable<string> GetValidatedColumns()
    {
        return _columnRules.Keys;
    }

    /// <summary>
    /// Má stĺpec nejaké validation rules
    /// </summary>
    public bool HasRules(string columnName)
    {
        return _columnRules.ContainsKey(columnName) && _columnRules[columnName].Count > 0;
    }

    /// <summary>
    /// Celkový počet validation rules
    /// </summary>
    public int TotalRulesCount => _columnRules.Values.Sum(rules => rules.Count);
}

/// <summary>
/// Jednotlivé validation rule pre stĺpec
/// </summary>
public class ValidationRule
{
    /// <summary>
    /// Názov rule (pre identifikáciu)
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Validation function (vracia true ak je hodnota valid)
    /// </summary>
    public Func<object?, bool> Validator { get; set; } = _ => true;

    /// <summary>
    /// Error message ak validation zlyhá
    /// </summary>
    public string ErrorMessage { get; set; } = "Validation failed";

    /// <summary>
    /// Priorita rule (vyššia hodnota = vyššia priorita)
    /// </summary>
    public int Priority { get; set; } = 0;

    /// <summary>
    /// Je rule zapnutá
    /// </summary>
    public bool IsEnabled { get; set; } = true;

    /// <summary>
    /// Validuje hodnotu
    /// </summary>
    public ValidationResult Validate(object? value)
    {
        if (!IsEnabled)
            return ValidationResult.Success();

        try
        {
            bool isValid = Validator(value);
            return isValid ? ValidationResult.Success() : ValidationResult.Error(ErrorMessage);
        }
        catch (Exception ex)
        {
            return ValidationResult.Error($"Validation error: {ex.Message}");
        }
    }
}

/// <summary>
/// Cross-row validation rule (validuje celé riadky/dataset)
/// </summary>
public class CrossRowValidationRule
{
    /// <summary>
    /// Názov rule
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Validation function pre celý dataset
    /// </summary>
    public Func<List<Dictionary<string, object?>>, CrossRowValidationResult> Validator { get; set; } = _ => CrossRowValidationResult.Success();

    /// <summary>
    /// Je rule zapnutá
    /// </summary>
    public bool IsEnabled { get; set; } = true;

    /// <summary>
    /// Validuje celý dataset
    /// </summary>
    public CrossRowValidationResult Validate(List<Dictionary<string, object?>> allRowData)
    {
        if (!IsEnabled)
            return CrossRowValidationResult.Success();

        try
        {
            return Validator(allRowData);
        }
        catch (Exception ex)
        {
            return CrossRowValidationResult.Error($"Cross-row validation error: {ex.Message}");
        }
    }
}

/// <summary>
/// Výsledok jednotlivej validation
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// Je validation úspešná
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Error message (ak validation zlyhala)
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Dodatočné info o validation
    /// </summary>
    public Dictionary<string, object?> AdditionalInfo { get; set; } = new();

    public static ValidationResult Success() => new() { IsValid = true };
    public static ValidationResult Error(string message) => new() { IsValid = false, ErrorMessage = message };
}

/// <summary>
/// Výsledok cross-row validation
/// </summary>
public class CrossRowValidationResult
{
    /// <summary>
    /// Je validation úspešná
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Error messages pre jednotlivé riadky (row index -> error message)
    /// </summary>
    public Dictionary<int, string> RowErrors { get; set; } = new();

    /// <summary>
    /// Global error message (applies to whole dataset)
    /// </summary>
    public string? GlobalErrorMessage { get; set; }

    public static CrossRowValidationResult Success() => new() { IsValid = true };
    public static CrossRowValidationResult Error(string globalMessage) => new() { IsValid = false, GlobalErrorMessage = globalMessage };
    public static CrossRowValidationResult ErrorWithRowDetails(Dictionary<int, string> rowErrors) => new() { IsValid = false, RowErrors = rowErrors };
}