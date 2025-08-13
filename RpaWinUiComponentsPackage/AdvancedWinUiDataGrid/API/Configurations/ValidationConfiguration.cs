namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid;

/// <summary>
/// Clean API configuration class pre validačné pravidlá DataGrid
/// Používa sa namiesto internal IValidationConfiguration
/// </summary>
public class ValidationConfiguration
{
    /// <summary>
    /// Validačné pravidlá pre jednotlivé stĺpce
    /// Key = názov stĺpca, Value = validačná funkcia
    /// </summary>
    public Dictionary<string, Func<object, bool>>? Rules { get; set; }
    
    /// <summary>
    /// Validačné pravidlá s custom chybovými správami
    /// Key = názov stĺpca, Value = (validačná funkcia, chybová správa)
    /// </summary>
    public Dictionary<string, (Func<object, bool> Validator, string ErrorMessage)>? RulesWithMessages { get; set; }
    
    /// <summary>Cross-row validačné pravidlá (validácia naprieč riadkami)</summary>
    public List<Func<List<Dictionary<string, object?>>, (bool IsValid, string? ErrorMessage)>>? CrossRowRules { get; set; }
    
    /// <summary>Je zapnutá real-time validácia (počas písania)</summary>
    public bool? EnableRealtimeValidation { get; set; }
    
    /// <summary>Je zapnutá batch validácia (validácia všetkých riadkov naraz)</summary>
    public bool? EnableBatchValidation { get; set; }
    
    /// <summary>Zobraziť validačné chyby vo ValidationAlerts stĺpci</summary>
    public bool? ShowValidationAlerts { get; set; }
    
    /// <summary>Zastaviť export/operácie ak sú validačné chyby</summary>
    public bool? StopOnValidationErrors { get; set; }
}