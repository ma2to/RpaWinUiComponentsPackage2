using Microsoft.Extensions.Logging.Abstractions;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Controls;

/// <summary>
/// AdvancedDataGrid - Table Management API
/// Partial class pre column names a row height management
/// Deleguje na AdvancedDataGridController a Table modul
/// </summary>
public sealed partial class AdvancedDataGrid
{
    #region Column Names API

    /// <summary>
    /// Získa všetky resolved column names (po duplicate resolution)
    /// Používa sa pre business logiku namiesto display names
    /// </summary>
    public List<string> GetResolvedColumnNames()
    {
        // Delegate to controller - business logic handled there
        return _controller.GetResolvedColumnNames();
    }

    /// <summary>
    /// Vráti info o tom či column name bol premenovaný during duplicate resolution
    /// </summary>
    public Dictionary<string, string> GetColumnNameMappings()
    {
        if (!IsInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // This would require storing the original->resolved mapping in SmartColumnNameResolver
        // For now, return empty mapping - can be enhanced later
        return new Dictionary<string, string>();
    }

    #endregion

    #region Row Height Management API

    // CurrentRowHeight property is already defined in main AdvancedDataGrid.cs

    /// <summary>
    /// Recalculate row height pre current dataset
    /// Volá sa automaticky pri import, ale môže sa volať manuálne
    /// </summary>
    public async Task RecalculateRowHeightAsync()
    {
        // Delegate to controller for business logic
        await _controller.RecalculateRowHeightAsync();
        
        // Update UI layout after row height change
        InvalidateLayout();
    }

    /// <summary>
    /// Set base row height (default pre prázdne bunky)
    /// </summary>
    public void SetBaseRowHeight(double baseHeight)
    {
        // Delegate to controller
        _controller.SetBaseRowHeight(baseHeight);
    }

    /// <summary>
    /// Enable/disable unlimited row height system
    /// </summary>
    public void SetUnlimitedRowHeightEnabled(bool enabled)
    {
        // Delegate to controller  
        _controller.SetUnlimitedRowHeightEnabled(enabled);
    }

    #endregion
}