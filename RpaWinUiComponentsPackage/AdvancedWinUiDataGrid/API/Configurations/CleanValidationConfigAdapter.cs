using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Validation.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid;

/// <summary>
/// Adapter pre konverziu clean ValidationConfiguration na internal IValidationConfiguration
/// </summary>
internal class CleanValidationConfigAdapter : IValidationConfiguration
{
    private readonly ValidationConfiguration _cleanConfig;

    public CleanValidationConfigAdapter(ValidationConfiguration cleanConfig)
    {
        _cleanConfig = cleanConfig;
    }

    public bool IsValidationEnabled => _cleanConfig.Rules?.Any() == true || _cleanConfig.RulesWithMessages?.Any() == true;

    public bool EnableRealtimeValidation => _cleanConfig.EnableRealtimeValidation ?? false;

    public bool EnableBatchValidation => _cleanConfig.EnableBatchValidation ?? false;

    public ValidationRuleSet GetValidationRules()
    {
        var ruleSet = new ValidationRuleSet();

        // Convert simple rules (Rules property)
        if (_cleanConfig.Rules != null)
        {
            foreach (var rule in _cleanConfig.Rules)
            {
                ruleSet.AddRule(rule.Key, new ValidationRule
                {
                    Name = $"{rule.Key}_Rule",
                    Validator = rule.Value,
                    ErrorMessage = $"{rule.Key} validation failed"
                });
            }
        }

        // Convert rules with custom messages (RulesWithMessages property)
        if (_cleanConfig.RulesWithMessages != null)
        {
            foreach (var rule in _cleanConfig.RulesWithMessages)
            {
                ruleSet.AddRule(rule.Key, new ValidationRule
                {
                    Name = $"{rule.Key}_RuleWithMessage",
                    Validator = rule.Value.Validator,
                    ErrorMessage = rule.Value.ErrorMessage
                });
            }
        }

        return ruleSet;
    }

    public List<CrossRowValidationRule> GetCrossRowValidationRules()
    {
        var crossRowRules = new List<CrossRowValidationRule>();

        if (_cleanConfig.CrossRowRules != null)
        {
            for (int i = 0; i < _cleanConfig.CrossRowRules.Count; i++)
            {
                var cleanRule = _cleanConfig.CrossRowRules[i];
                crossRowRules.Add(new CrossRowValidationRule
                {
                    Name = $"CrossRowRule_{i}",
                    Validator = allData =>
                    {
                        var result = cleanRule(allData);
                        return result.IsValid ? 
                            CrossRowValidationResult.Success() : 
                            CrossRowValidationResult.Error(result.ErrorMessage ?? "Cross-row validation failed");
                    }
                });
            }
        }

        return crossRowRules;
    }
}