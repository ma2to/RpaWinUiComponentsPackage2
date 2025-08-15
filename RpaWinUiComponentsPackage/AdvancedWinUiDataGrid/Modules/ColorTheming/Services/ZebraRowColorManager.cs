using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml.Media;
using Windows.UI;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Services;

/// <summary>
/// Zebra Row Color Manager - runtime color theming pre DataGrid rows
/// Implementuje zebra pattern (even/odd rows) s configurable colors from application
/// </summary>
public class ZebraRowColorManager
{
    #region Private Fields

    /// <summary>
    /// Current color configuration
    /// </summary>
    private DataGridColorConfig _colorConfig = DataGridColorConfig.Default;

    /// <summary>
    /// Logger (nullable)
    /// </summary>
    private readonly Microsoft.Extensions.Logging.ILogger? _logger;

    /// <summary>
    /// Je zebra row coloring enabled
    /// </summary>
    private bool _isZebraColoringEnabled = true;

    /// <summary>
    /// Cache pre pre-calculated brushes (performance optimization)
    /// </summary>
    private Dictionary<string, SolidColorBrush> _brushCache = new();

    #endregion

    #region Properties

    /// <summary>
    /// Current color configuration
    /// </summary>
    public DataGridColorConfig ColorConfig => _colorConfig;

    /// <summary>
    /// Je zebra coloring enabled
    /// </summary>
    public bool IsZebraColoringEnabled 
    { 
        get => _isZebraColoringEnabled; 
        set 
        { 
            _isZebraColoringEnabled = value; 
            _logger?.Info("🎨 ZEBRA COLORS: Zebra coloring {Status}", value ? "ENABLED" : "DISABLED");
        } 
    }

    #endregion

    #region Constructor

    /// <summary>
    /// Konštruktor s optional logger
    /// </summary>
    public ZebraRowColorManager(Microsoft.Extensions.Logging.ILogger? logger = null)
    {
        _logger = logger;
        _logger?.Info("🎨 ZEBRA COLOR MANAGER: Initialized");
    }

    #endregion

    #region Public API

    /// <summary>
    /// Aplikuje nové farby (runtime color theming) - už má garantované default values
    /// </summary>
    public void ApplyColorConfiguration(DataGridColorConfig colorConfig)
    {
        try
        {
            // POZNÁMKA: colorConfig už má garantované default hodnoty (nie sú null)
            _colorConfig = colorConfig ?? DataGridColorConfig.Default;
            
            // Clear brush cache (force recreation with new colors)
            _brushCache.Clear();
            
            _logger?.Info("🎨 ZEBRA COLORS: Applied new color configuration (no fallback needed - defaults guaranteed)");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: ApplyColorConfiguration failed");
        }
    }

    /// <summary>
    /// Získa background color pre row podľa row index (zebra pattern)
    /// POZNÁMKA: _colorConfig má už garantované default hodnoty po MergeWith()
    /// </summary>
    public SolidColorBrush GetRowBackgroundBrush(int rowIndex)
    {
        if (!_isZebraColoringEnabled)
        {
            return GetCachedBrush("default", _colorConfig.CellBackgroundColor!.Value);
        }

        try
        {
            bool isEvenRow = (rowIndex % 2) == 0;
            
            var color = isEvenRow 
                ? _colorConfig.EvenRowBackgroundColor!.Value
                : _colorConfig.OddRowBackgroundColor!.Value;

            var cacheKey = isEvenRow ? "even_row" : "odd_row";
            return GetCachedBrush(cacheKey, color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetRowBackgroundBrush failed - Row: {Row}", rowIndex);
            return GetCachedBrush("fallback", Microsoft.UI.Colors.White);
        }
    }

    /// <summary>
    /// Získa foreground color pre row
    /// POZNÁMKA: _colorConfig má už garantované default hodnoty po MergeWith()
    /// </summary>
    public SolidColorBrush GetRowForegroundBrush(int rowIndex)
    {
        try
        {
            var color = _colorConfig.CellForegroundColor!.Value;
            return GetCachedBrush("foreground", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetRowForegroundBrush failed - Row: {Row}", rowIndex);
            return GetCachedBrush("fallback_fg", Microsoft.UI.Colors.Black);
        }
    }

    /// <summary>
    /// Získa border color pre cell
    /// POZNÁMKA: _colorConfig má už garantované default hodnoty po MergeWith()
    /// </summary>
    public SolidColorBrush GetCellBorderBrush()
    {
        try
        {
            var color = _colorConfig.CellBorderColor!.Value; // Guaranteed čierny po MergeWith()
            return GetCachedBrush("border", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetCellBorderBrush failed");
            return GetCachedBrush("fallback_border", Microsoft.UI.Colors.Black);
        }
    }

    /// <summary>
    /// Získa selection background color (focus - bledo zelený)
    /// POZNÁMKA: _colorConfig má už garantované default hodnoty po MergeWith()
    /// </summary>
    public SolidColorBrush GetSelectionBackgroundBrush()
    {
        try
        {
            var color = _colorConfig.SelectionBackgroundColor!.Value; // Guaranteed bledo zelený po MergeWith()
            return GetCachedBrush("selection", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetSelectionBackgroundBrush failed");
            return GetCachedBrush("fallback_selection", Color.FromArgb(100, 144, 238, 144));
        }
    }

    /// <summary>
    /// Získa copy mode background color (bledo modrý)
    /// POZNÁMKA: _colorConfig má už garantované default hodnoty po MergeWith()
    /// </summary>
    public SolidColorBrush GetCopyModeBackgroundBrush()
    {
        try
        {
            var color = _colorConfig.CopyModeBackgroundColor!.Value; // Guaranteed bledo modrý po MergeWith()
            return GetCachedBrush("copy_mode", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetCopyModeBackgroundBrush failed");
            return GetCachedBrush("fallback_copy_mode", Color.FromArgb(100, 173, 216, 230));
        }
    }

    /// <summary>
    /// Získa hover background color
    /// </summary>
    public SolidColorBrush GetHoverBackgroundBrush()
    {
        try
        {
            var color = _colorConfig.HoverBackgroundColor ?? Color.FromArgb(50, 0, 0, 0);
            return GetCachedBrush("hover", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetHoverBackgroundBrush failed");
            return GetCachedBrush("fallback_hover", Color.FromArgb(50, 0, 0, 0));
        }
    }

    /// <summary>
    /// Získa validation error border color
    /// POZNÁMKA: _colorConfig má už garantované default hodnoty po MergeWith()
    /// </summary>
    public SolidColorBrush GetValidationErrorBorderBrush()
    {
        try
        {
            var color = _colorConfig.ValidationErrorBorderColor!.Value; // Guaranteed červený po MergeWith()
            return GetCachedBrush("validation_error_border", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetValidationErrorBorderBrush failed");
            return GetCachedBrush("fallback_validation", Microsoft.UI.Colors.Red);
        }
    }

    /// <summary>
    /// Získa header background color
    /// </summary>
    public SolidColorBrush GetHeaderBackgroundBrush()
    {
        try
        {
            var color = _colorConfig.HeaderBackgroundColor ?? Color.FromArgb(255, 240, 240, 240);
            return GetCachedBrush("header_bg", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetHeaderBackgroundBrush failed");
            return GetCachedBrush("fallback_header", Color.FromArgb(255, 240, 240, 240));
        }
    }

    /// <summary>
    /// Získa header foreground color
    /// </summary>
    public SolidColorBrush GetHeaderForegroundBrush()
    {
        try
        {
            var color = _colorConfig.HeaderForegroundColor ?? Microsoft.UI.Colors.Black;
            return GetCachedBrush("header_fg", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetHeaderForegroundBrush failed");
            return GetCachedBrush("fallback_header_fg", Microsoft.UI.Colors.Black);
        }
    }

    /// <summary>
    /// Získa header border color
    /// </summary>
    public SolidColorBrush GetHeaderBorderBrush()
    {
        try
        {
            var color = _colorConfig.HeaderBorderColor ?? Microsoft.UI.Colors.Black; // UPDATED: čierny header border
            return GetCachedBrush("header_border", color);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ZEBRA COLOR ERROR: GetHeaderBorderBrush failed");
            return GetCachedBrush("fallback_header_border", Microsoft.UI.Colors.Black); // UPDATED: čierny fallback
        }
    }

    /// <summary>
    /// Reset colors to default
    /// </summary>
    public void ResetToDefaults()
    {
        ApplyColorConfiguration(DataGridColorConfig.Default);
        _logger?.Info("🎨 ZEBRA COLORS: Reset to default colors");
    }

    /// <summary>
    /// Apply dark theme colors
    /// </summary>
    public void ApplyDarkTheme()
    {
        ApplyColorConfiguration(DataGridColorConfig.Dark);
        _logger?.Info("🎨 ZEBRA COLORS: Applied dark theme colors");
    }

    /// <summary>
    /// Clear brush cache (pre garbage collection)
    /// </summary>
    public void ClearCache()
    {
        _brushCache.Clear();
        _logger?.Info("🎨 ZEBRA COLORS: Brush cache cleared");
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Získa cached brush alebo vytvorí nový
    /// </summary>
    private SolidColorBrush GetCachedBrush(string cacheKey, Color color)
    {
        // Generate full cache key with color info
        var fullCacheKey = $"{cacheKey}_{color.A}_{color.R}_{color.G}_{color.B}";
        
        if (_brushCache.TryGetValue(fullCacheKey, out var existingBrush))
        {
            return existingBrush;
        }

        // Create new brush and cache it
        var newBrush = new SolidColorBrush(color);
        _brushCache[fullCacheKey] = newBrush;
        
        return newBrush;
    }

    #endregion

    #region Static Helpers

    /// <summary>
    /// Quick helper pre determining if row is even
    /// </summary>
    public static bool IsEvenRow(int rowIndex) => (rowIndex % 2) == 0;

    /// <summary>
    /// Quick helper pre determining if row is odd
    /// </summary>
    public static bool IsOddRow(int rowIndex) => (rowIndex % 2) != 0;

    #endregion
}