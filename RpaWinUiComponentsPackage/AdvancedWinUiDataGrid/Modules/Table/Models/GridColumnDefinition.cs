namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;

/// <summary>
/// Definícia stĺpca pre DataGrid
/// </summary>
public class GridColumnDefinition
{
    /// <summary>
    /// Názov stĺpca (property name)
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Zobrazovaný názov stĺpca (header text)
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Dátový typ stĺpca
    /// </summary>
    public Type DataType { get; set; } = typeof(string);

    /// <summary>
    /// Šírka stĺpca (null = auto)
    /// </summary>
    public double? Width { get; set; }

    /// <summary>
    /// Minimálna šírka stĺpca
    /// </summary>
    public double MinWidth { get; set; } = 50;

    /// <summary>
    /// Maximálna šírka stĺpca (null = unlimited)
    /// </summary>
    public double? MaxWidth { get; set; }

    /// <summary>
    /// Je stĺpec viditeľný
    /// </summary>
    public bool IsVisible { get; set; } = true;

    /// <summary>
    /// Je stĺpec editovateľný
    /// </summary>
    public bool IsReadOnly { get; set; } = false;

    /// <summary>
    /// Je stĺpec sortovateľný
    /// </summary>
    public bool IsSortable { get; set; } = true;

    /// <summary>
    /// Je stĺpec filtrovateľný
    /// </summary>
    public bool IsFilterable { get; set; } = true;

    /// <summary>
    /// Je to CheckBox stĺpec (special column)
    /// </summary>
    public bool IsCheckBoxColumn { get; set; } = false;

    /// <summary>
    /// Je to ValidationAlerts stĺpec (special column - automaticky second-to-last position)
    /// </summary>
    public bool IsValidationAlertsColumn { get; set; } = false;

    /// <summary>
    /// Je to DeleteRow stĺpec (special column - automaticky last position)
    /// </summary>
    public bool IsDeleteRowColumn { get; set; } = false;

    /// <summary>
    /// Custom validation rules pre stĺpec
    /// </summary>
    public List<string> ValidationRules { get; set; } = new();

    /// <summary>
    /// Tooltip text pre header
    /// </summary>
    public string? ToolTip { get; set; }

    /// <summary>
    /// Format string pre zobrazenie hodnôt (DateTime, Number formatting)
    /// </summary>
    public string? Format { get; set; }

    /// <summary>
    /// Default hodnota pre nové bunky
    /// </summary>
    public object? DefaultValue { get; set; }

    /// <summary>
    /// Text alignment v stĺpci
    /// </summary>
    public ColumnAlignment TextAlignment { get; set; } = ColumnAlignment.Left;

    /// <summary>
    /// Je to special column (CheckBox, ValidationAlerts, DeleteRow)
    /// </summary>
    public bool IsSpecialColumn => IsCheckBoxColumn || IsValidationAlertsColumn || IsDeleteRowColumn;

    /// <summary>
    /// Klonuje column definition
    /// </summary>
    public GridColumnDefinition Clone()
    {
        return new GridColumnDefinition
        {
            Name = Name,
            DisplayName = DisplayName,
            DataType = DataType,
            Width = Width,
            MinWidth = MinWidth,
            MaxWidth = MaxWidth,
            IsVisible = IsVisible,
            IsReadOnly = IsReadOnly,
            IsSortable = IsSortable,
            IsFilterable = IsFilterable,
            IsCheckBoxColumn = IsCheckBoxColumn,
            IsValidationAlertsColumn = IsValidationAlertsColumn,
            IsDeleteRowColumn = IsDeleteRowColumn,
            ValidationRules = new List<string>(ValidationRules),
            ToolTip = ToolTip,
            Format = Format,
            DefaultValue = DefaultValue,
            TextAlignment = TextAlignment
        };
    }

    /// <summary>
    /// Validuje column definition
    /// </summary>
    public bool IsValid(out string? errorMessage)
    {
        if (string.IsNullOrEmpty(Name))
        {
            errorMessage = "Column Name cannot be empty";
            return false;
        }

        if (string.IsNullOrEmpty(DisplayName))
        {
            errorMessage = "Column DisplayName cannot be empty";
            return false;
        }

        if (MinWidth < 0)
        {
            errorMessage = "MinWidth must be >= 0";
            return false;
        }

        if (MaxWidth.HasValue && MaxWidth.Value < MinWidth)
        {
            errorMessage = "MaxWidth must be >= MinWidth";
            return false;
        }

        if (Width.HasValue && (Width.Value < MinWidth || (MaxWidth.HasValue && Width.Value > MaxWidth.Value)))
        {
            errorMessage = "Width must be between MinWidth and MaxWidth";
            return false;
        }

        // Special columns validation
        int specialColumnCount = 0;
        if (IsCheckBoxColumn) specialColumnCount++;
        if (IsValidationAlertsColumn) specialColumnCount++;
        if (IsDeleteRowColumn) specialColumnCount++;

        if (specialColumnCount > 1)
        {
            errorMessage = "Column cannot be multiple special column types";
            return false;
        }

        errorMessage = null;
        return true;
    }

    public override string ToString()
    {
        var type = IsSpecialColumn ? $" [{(IsCheckBoxColumn ? "CheckBox" : IsValidationAlertsColumn ? "ValidationAlerts" : "DeleteRow")}]" : "";
        return $"{DisplayName} ({Name}){type}";
    }
}

/// <summary>
/// Text alignment pre stĺpce
/// </summary>
public enum ColumnAlignment
{
    Left,
    Center,
    Right
}