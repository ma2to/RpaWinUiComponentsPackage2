using Microsoft.Extensions.Logging.Abstractions;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Services;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Controls;

/// <summary>
/// AdvancedDataGrid - Color Configuration API
/// Partial class pre color management funkcionalitu
/// Deleguje na AdvancedDataGridController a ColorTheming modul
/// </summary>
public sealed partial class AdvancedDataGrid
{
    #region Color Configuration API

    /// <summary>
    /// Aplikuje nové farby okamžite - SELECTIVE MERGE (len nastavené farby, zvyšok zostane default)
    /// Ak aplikácia nenastaví farbu (null), použije sa default farba
    /// Ak aplikácia nenastaví žiadne farby, všetko zostane default
    /// </summary>
    public void ApplyColorConfig(DataGridColorConfig? colorConfig = null)
    {
        // Delegate to controller for business logic
        _controller.ApplyColorConfig(colorConfig);
        
        // Update UI layer with new colors from controller
        UpdateXAMLProperties(_controller.ColorConfig);
        ApplyColorConfiguration(_controller.ColorConfig);
    }

    /// <summary>
    /// Resetuje farby na default
    /// </summary>
    public void ResetColorsToDefaults()
    {
        ApplyColorConfig(DataGridColorConfig.Default);
    }

    /// <summary>
    /// Aplikuje dark theme colors
    /// </summary>
    public void ApplyDarkTheme()
    {
        ApplyColorConfig(DataGridColorConfig.Dark);
        _logger?.Info("🎨 COLOR CONFIG: Applied dark theme");
    }

    /// <summary>
    /// Enable/disable zebra row coloring
    /// </summary>
    public void SetZebraColoringEnabled(bool enabled)
    {
        // Delegate to controller
        _controller.SetZebraColoringEnabled(enabled);
    }

    // IsZebraColoringEnabled property is already defined in main AdvancedDataGrid.cs

    /// <summary>
    /// Získa zebra color manager (pre advanced scenarios)
    /// </summary>
    public ZebraRowColorManager? ZebraColorManager => _controller.ZebraColorManager;

    #endregion
}