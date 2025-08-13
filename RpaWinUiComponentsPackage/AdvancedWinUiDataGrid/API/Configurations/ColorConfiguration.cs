namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid;

/// <summary>
/// Clean API configuration class pre farby DataGrid elementov
/// Používa sa namiesto internal DataGridColorConfig
/// Všetky farby sú voliteľné - ak nie sú nastavené, použijú sa default hodnoty
/// </summary>
public class ColorConfiguration
{
    /// <summary>Farba pozadia buniek</summary>
    public string? CellBackground { get; set; }
    
    /// <summary>Farba textu v bunkách</summary>
    public string? CellForeground { get; set; }
    
    /// <summary>Farba okrajov buniek</summary>
    public string? CellBorder { get; set; }
    
    /// <summary>Farba pozadia hlavičky</summary>
    public string? HeaderBackground { get; set; }
    
    /// <summary>Farba textu hlavičky</summary>
    public string? HeaderForeground { get; set; }
    
    /// <summary>Farba okrajov hlavičky</summary>
    public string? HeaderBorder { get; set; }
    
    /// <summary>Farba pozadia pri označení/selection</summary>
    public string? SelectionBackground { get; set; }
    
    /// <summary>Farba textu pri označení/selection</summary>
    public string? SelectionForeground { get; set; }
    
    /// <summary>Farba okraja validačných chýb</summary>
    public string? ValidationErrorBorder { get; set; }
    
    /// <summary>Farba pozadia validačných chýb</summary>
    public string? ValidationErrorBackground { get; set; }
    
    /// <summary>Farba pozadia párnych riadkov (zebra efekt)</summary>
    public string? AlternateRowBackground { get; set; }
    
    /// <summary>Farba pozadia pri hover nad riadkom</summary>
    public string? HoverBackground { get; set; }
    
    /// <summary>Farba pozadia status baru</summary>
    public string? StatusBarBackground { get; set; }
    
    /// <summary>Farba textu status baru</summary>
    public string? StatusBarForeground { get; set; }
}