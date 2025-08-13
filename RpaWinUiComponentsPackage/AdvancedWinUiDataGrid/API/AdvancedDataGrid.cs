using Microsoft.Extensions.Logging;
using System.Linq;
using System.Data;
using InternalControl = RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Controls.AdvancedDataGrid;
using InternalGridColumnDefinition = RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models.GridColumnDefinition;
using InternalValidationConfig = RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Validation.Models.IValidationConfiguration;
using InternalThrottlingConfig = RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models.GridThrottlingConfig;
using InternalColorConfig = RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models.DataGridColorConfig;
using InternalBatchValidationResult = RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Services.BatchValidationResult;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid;

/// <summary>
/// CLEAN PUBLIC API wrapper pre AdvancedDataGrid control
/// External aplikácie používajú tento clean API wrapper namiesto internal control
/// </summary>
public class AdvancedDataGrid : Microsoft.UI.Xaml.Controls.UserControl
{
    private readonly InternalControl _internalControl;

    public AdvancedDataGrid()
    {
        _internalControl = new InternalControl();
        this.Content = _internalControl;
    }

    /// <summary>
    /// CLEAN API InitializeAsync - uses clean Configuration classes
    /// </summary>
    public async Task InitializeAsync(
        List<ColumnConfiguration> columns,
        ColorConfiguration? colors = null,
        ValidationConfiguration? validation = null,
        PerformanceConfiguration? performance = null,
        int emptyRowsCount = 15,
        ILogger? logger = null,
        bool enableSort = false,
        bool enableSearch = false,
        bool enableFilter = false,
        double? minWidth = null,
        double? minHeight = null,
        double? maxWidth = null,
        double? maxHeight = null)
    {
        // Convert clean Configuration classes to internal types
        var internalColumns = ConvertColumnsToInternal(columns);
        var internalColorConfig = ConvertColorsToInternal(colors);
        var internalValidationConfig = ConvertValidationToInternal(validation);
        var internalPerformanceConfig = ConvertPerformanceToInternal(performance);

        // Call internal control
        await _internalControl.InitializeAsync(
            internalColumns, internalValidationConfig, internalPerformanceConfig, 
            emptyRowsCount, internalColorConfig, logger, 
            validation?.EnableBatchValidation ?? false,
            performance?.MaxSearchHistoryItems ?? 0, 
            enableSort, enableSearch, enableFilter, 0,
            minWidth, minHeight, maxWidth, maxHeight);
    }

    /// <summary>
    /// Import from Dictionary - clean API
    /// </summary>
    public async Task ImportFromDictionaryAsync(List<Dictionary<string, object?>> data) =>
        await _internalControl.ImportFromDictionaryAsync(data);

    /// <summary>
    /// Import from DataTable - clean API
    /// </summary>
    public async Task ImportFromDataTableAsync(DataTable dataTable) =>
        await _internalControl.ImportFromDataTableAsync(dataTable);

    /// <summary>
    /// Export to Dictionary - clean API
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ExportToDictionaryAsync(bool includeValidAlerts = false, bool removeAfter = false) =>
        await _internalControl.ExportToDictionaryAsync(includeValidAlerts, removeAfter);

    /// <summary>
    /// Export to DataTable - clean API
    /// </summary>
    public async Task<DataTable> ExportToDataTableAsync(bool includeValidAlerts = false, bool removeAfter = false) =>
        await _internalControl.ExportToDataTableAsync(includeValidAlerts, removeAfter);

    /// <summary>
    /// Smart Delete Row - clean API
    /// </summary>
    public async Task SmartDeleteRowAsync(int rowIndex) =>
        await _internalControl.SmartDeleteRowAsync(rowIndex);

    /// <summary>
    /// Smart Delete Multiple Rows - clean API
    /// </summary>
    public async Task SmartDeleteRowAsync(List<int> rowIndices) =>
        await _internalControl.SmartDeleteRowAsync(rowIndices);

    /// <summary>
    /// Clear all data - clean API
    /// </summary>
    public async Task ClearAllDataAsync() =>
        await _internalControl.ClearAllDataAsync();

    /// <summary>
    /// Refresh UI - clean API
    /// </summary>
    public async Task RefreshUIAsync() =>
        await _internalControl.RefreshUIAsync();

    /// <summary>
    /// Validation methods - clean API
    /// </summary>
    public async Task<bool> AreAllNonEmptyRowsValidAsync() =>
        await _internalControl.AreAllNonEmptyRowsValidAsync();

    public async Task<InternalBatchValidationResult?> ValidateAllRowsBatchAsync() =>
        await _internalControl.ValidateAllRowsBatchAsync();

    /// <summary>
    /// Statistics - clean API
    /// </summary>
    public int GetTotalRowCount() => _internalControl.GetTotalRowCount();
    public int GetColumnCount() => _internalControl.GetColumnCount();
    public async Task<int> GetVisibleRowsCountAsync() => await _internalControl.GetVisibleRowsCountAsync();
    public bool HasData => _internalControl.HasData;
    public async Task<int> GetLastDataRowAsync() => await _internalControl.GetLastDataRowAsync();
    public int GetMinimumRowCount() => _internalControl.GetMinimumRowCount();

    /// <summary>
    /// UI operations - clean API
    /// </summary>
    public async Task UpdateValidationUIAsync() => await _internalControl.UpdateValidationUIAsync();
    public void InvalidateLayout() => _internalControl.InvalidateLayout();
    public async Task CompactRowsAsync() => await _internalControl.CompactRowsAsync();
    
    /// <summary>
    /// Color operations - clean API
    /// </summary>
    public void ApplyColorConfig(InternalColorConfig colorConfig) => 
        _internalControl.ApplyColorConfig(colorConfig ?? InternalColorConfig.Default);
    public void ResetColorsToDefaults() => _internalControl.ResetColorsToDefaults();

    /// <summary>
    /// Paste operations - clean API
    /// </summary>
    public async Task PasteDataAsync(List<Dictionary<string, object?>> pasteData, int startRow, int startColumn) =>
        await _internalControl.PasteDataAsync(pasteData, startRow, startColumn);

    #region Clean API to Internal Type Conversions

    /// <summary>
    /// Convert clean ColumnConfiguration to internal GridColumnDefinition
    /// </summary>
    private List<InternalGridColumnDefinition> ConvertColumnsToInternal(List<ColumnConfiguration> columns)
    {
        return columns.Select(col => new InternalGridColumnDefinition
        {
            Name = col.Name ?? string.Empty,
            DataType = col.Type ?? typeof(object),
            Width = col.Width ?? 100,
            DisplayName = col.DisplayName ?? col.Name ?? string.Empty,
            IsValidationAlertsColumn = col.IsValidationColumn ?? false,
            IsDeleteRowColumn = col.IsDeleteColumn ?? false,
            IsCheckBoxColumn = col.IsCheckboxColumn ?? false,  // Correct property name
            MinWidth = col.MinWidth ?? 50,
            MaxWidth = col.MaxWidth ?? 500,
            IsReadOnly = col.IsReadOnly ?? false,
            IsVisible = col.IsVisible ?? true
            // Note: Order removed - doesn't exist in internal type
        }).ToList();
    }

    /// <summary>
    /// Convert clean ColorConfiguration to internal DataGridColorConfig
    /// </summary>
    private InternalColorConfig ConvertColorsToInternal(ColorConfiguration? colors)
    {
        if (colors == null) return InternalColorConfig.Default;

        var config = new InternalColorConfig();
        
        if (!string.IsNullOrEmpty(colors.CellBackground))
            config.CellBackgroundColor = ParseColor(colors.CellBackground);
        if (!string.IsNullOrEmpty(colors.CellForeground))
            config.CellForegroundColor = ParseColor(colors.CellForeground);
        if (!string.IsNullOrEmpty(colors.CellBorder))
            config.CellBorderColor = ParseColor(colors.CellBorder);
        if (!string.IsNullOrEmpty(colors.HeaderBackground))
            config.HeaderBackgroundColor = ParseColor(colors.HeaderBackground);
        if (!string.IsNullOrEmpty(colors.HeaderForeground))
            config.HeaderForegroundColor = ParseColor(colors.HeaderForeground);
        if (!string.IsNullOrEmpty(colors.HeaderBorder))
            config.HeaderBorderColor = ParseColor(colors.HeaderBorder);
        if (!string.IsNullOrEmpty(colors.SelectionBackground))
            config.SelectionBackgroundColor = ParseColor(colors.SelectionBackground);
        if (!string.IsNullOrEmpty(colors.SelectionForeground))
            config.SelectionForegroundColor = ParseColor(colors.SelectionForeground);
        if (!string.IsNullOrEmpty(colors.ValidationErrorBorder))
            config.ValidationErrorBorderColor = ParseColor(colors.ValidationErrorBorder);
        if (!string.IsNullOrEmpty(colors.ValidationErrorBackground))
            config.ValidationErrorBackgroundColor = ParseColor(colors.ValidationErrorBackground);

        return config;
    }

    /// <summary>
    /// Convert clean ValidationConfiguration to internal IValidationConfiguration
    /// </summary>
    private InternalValidationConfig? ConvertValidationToInternal(ValidationConfiguration? validation)
    {
        if (validation == null) return null;

        // Create internal validation config implementation
        return new CleanValidationConfigAdapter(validation);
    }

    /// <summary>
    /// Convert clean PerformanceConfiguration to internal GridThrottlingConfig
    /// </summary>
    private InternalThrottlingConfig ConvertPerformanceToInternal(PerformanceConfiguration? performance)
    {
        if (performance == null) return InternalThrottlingConfig.Default;

        return new InternalThrottlingConfig
        {
            VirtualizationBufferSize = performance.VirtualizationThreshold ?? 1000,
            BulkOperationBatchSize = performance.BatchSize ?? 100,
            UIUpdateIntervalMs = performance.RenderDelayMs ?? 50,
            SearchUpdateIntervalMs = performance.SearchThrottleMs ?? 300,
            ValidationUpdateIntervalMs = performance.ValidationThrottleMs ?? 500,
            MaxSearchHistoryItems = performance.MaxSearchHistoryItems ?? 100,
            EnableBackgroundProcessing = performance.EnableUIThrottling ?? true,
            EnableAggressiveMemoryManagement = performance.EnableLazyLoading ?? false
        };
    }

    /// <summary>
    /// Parse color string to Windows.UI.Color
    /// </summary>
    private Windows.UI.Color ParseColor(string colorString)
    {
        try
        {
            // Remove # if present
            colorString = colorString.TrimStart('#');
            
            // Parse hex color
            if (colorString.Length == 6)
            {
                var r = Convert.ToByte(colorString.Substring(0, 2), 16);
                var g = Convert.ToByte(colorString.Substring(2, 2), 16);
                var b = Convert.ToByte(colorString.Substring(4, 2), 16);
                return Windows.UI.Color.FromArgb(255, r, g, b);
            }
            else if (colorString.Length == 8)
            {
                var a = Convert.ToByte(colorString.Substring(0, 2), 16);
                var r = Convert.ToByte(colorString.Substring(2, 2), 16);
                var g = Convert.ToByte(colorString.Substring(4, 2), 16);
                var b = Convert.ToByte(colorString.Substring(6, 2), 16);
                return Windows.UI.Color.FromArgb(a, r, g, b);
            }
        }
        catch
        {
            // Return default color if parsing fails
        }
        
        return Microsoft.UI.Colors.Transparent;
    }

    #endregion
}