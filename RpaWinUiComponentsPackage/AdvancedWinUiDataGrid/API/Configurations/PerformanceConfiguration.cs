namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid;

/// <summary>
/// Clean API configuration class pre výkonnostné nastavenia DataGrid
/// Používa sa namiesto internal GridThrottlingConfig
/// </summary>
public class PerformanceConfiguration
{
    /// <summary>Počet riadkov od ktorého sa zapne virtualizácia</summary>
    public int? VirtualizationThreshold { get; set; }
    
    /// <summary>Veľkosť dávky pri batch operáciách</summary>
    public int? BatchSize { get; set; }
    
    /// <summary>Oneskorenie pri renderovaní UI (milliseconds)</summary>
    public int? RenderDelayMs { get; set; }
    
    /// <summary>Throttling delay pre search operácie (milliseconds)</summary>
    public int? SearchThrottleMs { get; set; }
    
    /// <summary>Throttling delay pre validation operácie (milliseconds)</summary>
    public int? ValidationThrottleMs { get; set; }
    
    /// <summary>Maximálny počet search history položiek</summary>
    public int? MaxSearchHistoryItems { get; set; }
    
    /// <summary>Je zapnuté UI throttling pre lepšiu responzivitu</summary>
    public bool? EnableUIThrottling { get; set; }
    
    /// <summary>Je zapnutá lazy loading pre veľké datasety</summary>
    public bool? EnableLazyLoading { get; set; }
}