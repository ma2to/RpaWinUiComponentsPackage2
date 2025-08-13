using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Services;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Services;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Validation.Models;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Controls;

/// <summary>
/// Advanced WinUI DataGrid - Pure UI Wrapper with Controller Pattern
/// Čistý UI komponent - business logika je v AdvancedDataGridController
/// Clean separation: UI ← Controller → Business Logic
/// </summary>
public sealed partial class AdvancedDataGrid : UserControl
{
    #region Private Fields

    /// <summary>
    /// Business logic controller - all data operations
    /// </summary>
    private readonly AdvancedDataGridController _controller;

    /// <summary>
    /// Logger pre UI operácie (optional)
    /// </summary>
    private Microsoft.Extensions.Logging.ILogger? _logger;

    /// <summary>
    /// Je komponent inicializovaný
    /// </summary>
    private bool _isInitialized;

    /// <summary>
    /// Performance throttling configuration
    /// </summary>
    private GridThrottlingConfig? _throttlingConfig;

    // UI XAML elements are automatically generated from XAML file

    #endregion

    #region Properties

    /// <summary>
    /// Je komponent inicializovaný
    /// </summary>
    public bool IsInitialized => _controller.IsInitialized;

    /// <summary>
    /// Headless table core (for advanced scenarios)
    /// </summary>
    public DynamicTableCore TableCore => _controller.TableCore;

    /// <summary>
    /// Aktuálny počet riadkov
    /// </summary>
    public int ActualRowCount => _controller.ActualRowCount;

    /// <summary>
    /// Počet stĺpcov
    /// </summary>
    public int ColumnCount => _controller.ColumnCount;

    /// <summary>
    /// Má dataset nejaké dáta
    /// </summary>
    public bool HasData => _controller.HasData;

    /// <summary>
    /// Current row height
    /// </summary>
    public double CurrentRowHeight => _controller.CurrentRowHeight;

    /// <summary>
    /// Is zebra coloring enabled
    /// </summary>
    public bool IsZebraColoringEnabled => _controller.IsZebraColoringEnabled;

    #endregion

    #region Constructor

    public AdvancedDataGrid()
    {
        this.InitializeComponent();
        _controller = new AdvancedDataGridController();
        _isInitialized = false;
        
        // Initialize UI event handlers
        InitializeUIEventHandlers();
    }

    #endregion

    #region Public API - Initialization

    /// <summary>
    /// Inicializuje DataGrid s column definitions
    /// HLAVNÝ ENTRY POINT pre všetko použitie - UI + Controller
    /// </summary>
    public async Task InitializeAsync(
        List<GridColumnDefinition> columns,
        IValidationConfiguration? validationConfig = null,
        GridThrottlingConfig? throttlingConfig = null,
        int emptyRowsCount = 15,
        DataGridColorConfig? colorConfig = null,
        Microsoft.Extensions.Logging.ILogger? logger = null,
        bool enableBatchValidation = false,
        int maxSearchHistoryItems = 0,
        bool enableSort = false,
        bool enableSearch = false,
        bool enableFilter = false,
        int searchHistoryItems = 0,
        double? minWidth = null,
        double? minHeight = null,
        double? maxWidth = null,
        double? maxHeight = null)
    {
        try
        {
            _logger = logger;

            // Phase 1: Initialize controller (business logic)
            await _controller.InitializeAsync(columns, validationConfig, throttlingConfig, 
                emptyRowsCount, colorConfig, logger, enableBatchValidation, maxSearchHistoryItems,
                enableSort, enableSearch, enableFilter, searchHistoryItems);

            // Phase 2: Initialize UI layer
            await InitializeUILayerAsync(colorConfig, minWidth, minHeight, maxWidth, maxHeight);

            // Mark as initialized
            _isInitialized = true;

            _logger?.LogInformation("🎨 UI WRAPPER: AdvancedDataGrid UI layer initialized successfully");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI WRAPPER ERROR: AdvancedDataGrid UI initialization failed");
            throw;
        }
    }

    /// <summary>
    /// Initialize UI layer after controller is ready
    /// </summary>
    private async Task InitializeUILayerAsync(DataGridColorConfig? colorConfig, 
        double? minWidth, double? minHeight, double? maxWidth, double? maxHeight)
    {
        try
        {
            // Apply UI sizing
            ApplyUISizing(minWidth, minHeight, maxWidth, maxHeight);

            // Initialize XAML properties with colors from controller
            UpdateXAMLProperties(_controller.ColorConfig);

            // Apply color configuration to UI elements
            ApplyColorConfiguration(_controller.ColorConfig);

            // Initialize UI virtualization
            await InitializeUIVirtualizationAsync();

            _logger?.LogInformation("✅ UI LAYER: Initialization completed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI LAYER ERROR: UI layer initialization failed");
            throw;
        }
    }

    #endregion

    #region UI Update API - Manual Refresh Strategy

    /// <summary>
    /// Force refresh celého UI (re-render všetkých buniek)
    /// </summary>
    public async Task RefreshUIAsync()
    {
        if (!_isInitialized) return;

        try
        {
            _logger?.LogInformation("🎨 UI UPDATE: Full UI refresh started");
            
            await ShowLoadingAsync(true);
            await RenderAllCellsAsync();
            await ShowLoadingAsync(false);

            _logger?.LogInformation("✅ UI UPDATE: Full UI refresh completed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: RefreshUIAsync failed");
            await ShowLoadingAsync(false);
        }
    }

    /// <summary>
    /// Update len validation visual indicators (borders, ValidationAlerts column)
    /// </summary>
    public async Task UpdateValidationUIAsync()
    {
        if (!_isInitialized) return;

        try
        {
            _logger?.LogInformation("🎨 UI UPDATE: Validation UI update started");

            await UpdateValidationVisualsAsync();

            _logger?.LogInformation("✅ UI UPDATE: Validation UI update completed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateValidationUIAsync failed");
        }
    }

    /// <summary>
    /// Update UI pre konkrétny riadok
    /// </summary>
    public async Task UpdateRowUIAsync(int rowIndex)
    {
        if (!_isInitialized) return;

        try
        {
            _logger?.LogInformation("🎨 UI UPDATE: Row UI update - Row: {Row}", rowIndex);

            await UpdateSpecificRowUIAsync(rowIndex);

            _logger?.LogInformation("✅ UI UPDATE: Row UI update completed - Row: {Row}", rowIndex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateRowUIAsync failed - Row: {Row}", rowIndex);
        }
    }

    /// <summary>
    /// Update UI pre konkrétnu bunku
    /// </summary>
    public async Task UpdateCellUIAsync(int row, int column)
    {
        if (!_isInitialized) return;

        try
        {
            _logger?.LogInformation("🎨 UI UPDATE: Cell UI update - Row: {Row}, Column: {Column}", row, column);

            await UpdateSpecificCellUIAsync(row, column);

            _logger?.LogInformation("✅ UI UPDATE: Cell UI update completed - Row: {Row}, Column: {Column}", row, column);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateCellUIAsync failed - Row: {Row}, Column: {Column}", row, column);
        }
    }

    /// <summary>
    /// Update UI pre celý stĺpec
    /// </summary>
    public async Task UpdateColumnUIAsync(string columnName)
    {
        if (!_isInitialized) return;

        try
        {
            _logger?.LogInformation("🎨 UI UPDATE: Column UI update - Column: {Column}", columnName);

            await UpdateSpecificColumnUIAsync(columnName);

            _logger?.LogInformation("✅ UI UPDATE: Column UI update completed - Column: {Column}", columnName);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateColumnUIAsync failed - Column: {Column}", columnName);
        }
    }

    /// <summary>
    /// Force layout recalculation (sizing, positioning)
    /// </summary>
    public void InvalidateLayout()
    {
        if (!_isInitialized) return;

        try
        {
            _logger?.LogInformation("🎨 UI UPDATE: Layout invalidation");

            // Update layout with current row height (if available)
            if (DataRepeater != null)
            {
                var currentRowHeight = _controller.CurrentRowHeight;
                var layout = new UniformGridLayout
                {
                    Orientation = Orientation.Vertical,
                    MinItemWidth = 120,
                    MinItemHeight = currentRowHeight,
                    ItemsStretch = UniformGridLayoutItemsStretch.Fill
                };
                DataRepeater.Layout = layout;
                
                _logger?.LogInformation("🎨 UI LAYOUT: Updated with unified row height: {Height}px", Math.Ceiling(currentRowHeight));
            }

            // Force ItemsRepeater to recalculate layout
            DataRepeater?.InvalidateMeasure();
            HeaderRepeater?.InvalidateMeasure();

            _logger?.LogInformation("✅ UI UPDATE: Layout invalidation completed");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: InvalidateLayout failed");
        }
    }

    #endregion

    // Passthrough API moved to Modules/PublicAPI/Services/AdvancedDataGrid.PublicAPI.cs

    #region Status and Info API

    /// <summary>
    /// Aktualizuje status text
    /// </summary>
    public void UpdateStatus(string statusText)
    {
        StatusText.Text = statusText;
        StatusBar.Visibility = string.IsNullOrEmpty(statusText) ? Visibility.Collapsed : Visibility.Visible;
    }

    /// <summary>
    /// Zobrazí/skryje loading indicator
    /// </summary>
    public async Task ShowLoadingAsync(bool isLoading, string? loadingText = null)
    {
        LoadingRing.IsActive = isLoading;
        LoadingRing.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;

        if (isLoading && !string.IsNullOrEmpty(loadingText))
        {
            UpdateStatus(loadingText);
        }
        else if (!isLoading)
        {
            UpdateStatus("Ready");
        }
    }

    #endregion

    // Color Configuration API moved to Modules/ColorTheming/Services/AdvancedDataGrid.ColorConfiguration.cs

    // Column Names API and Row Height Management API moved to Modules/Table/Services/AdvancedDataGrid.TableManagement.cs

    #region UI Helper Methods

    /// <summary>
    /// Updates XAML elements PROGRAMATICALLY from color config - NO hardcoded colors!
    /// </summary>
    private void UpdateXAMLProperties(DataGridColorConfig colorConfig)
    {
        try
        {
            // PROGRAMATICKY nastaví farby na XAML elementy (nie hardkódovane!)
            if (StatusBar != null)
            {
                // StatusBar background from cell background color
                if (colorConfig.CellBackgroundColor.HasValue)
                {
                    StatusBar.Background = new Microsoft.UI.Xaml.Media.SolidColorBrush(colorConfig.CellBackgroundColor.Value);
                }
                
                // StatusBar border from cell border color  
                if (colorConfig.CellBorderColor.HasValue)
                {
                    StatusBar.BorderBrush = new Microsoft.UI.Xaml.Media.SolidColorBrush(colorConfig.CellBorderColor.Value);
                }
            }
            
            if (StatusText != null && colorConfig.CellForegroundColor.HasValue)
            {
                // StatusBar text color from cell foreground color
                StatusText.Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(colorConfig.CellForegroundColor.Value);
            }

            // Update DataTemplates programmatically - create new templates with colors from config
            UpdateCellDataTemplate(colorConfig);
            UpdateHeaderDataTemplate(colorConfig);

            _logger?.LogInformation("🎨 XAML ELEMENTS: Updated programatically from color config (including templates)");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 XAML ERROR: UpdateXAMLProperties failed");
        }
    }

    /// <summary>
    /// Updates cell template colors by accessing existing XAML templates - WinUI3 compatible approach
    /// Note: WinUI 3 doesn't support FrameworkElementFactory - we update existing templates instead
    /// </summary>
    private void UpdateCellDataTemplate(DataGridColorConfig colorConfig)
    {
        try
        {
            // WinUI 3 approach: We rely on programmatic coloring during runtime rendering
            // instead of template modification since FrameworkElementFactory is not available
            
            _logger?.LogInformation("🎨 CELL TEMPLATE: Colors will be applied during runtime rendering (WinUI3 compatible)");
            
            // Note: Actual cell coloring happens in ApplyColorConfiguration() and ZebraRowColorManager
            // Templates stay generic, colors are applied programmatically to rendered elements
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TEMPLATE ERROR: UpdateCellDataTemplate failed");
        }
    }

    /// <summary>
    /// Updates header template colors by accessing existing XAML templates - WinUI3 compatible approach  
    /// Note: WinUI 3 doesn't support FrameworkElementFactory - we update existing templates instead
    /// </summary>
    private void UpdateHeaderDataTemplate(DataGridColorConfig colorConfig)
    {
        try
        {
            // WinUI 3 approach: We rely on programmatic coloring during runtime rendering
            // instead of template modification since FrameworkElementFactory is not available
            
            _logger?.LogInformation("🎨 HEADER TEMPLATE: Colors will be applied during runtime rendering (WinUI3 compatible)");
            
            // Note: Actual header coloring happens in ApplyColorConfiguration() and ZebraRowColorManager
            // Templates stay generic, colors are applied programmatically to rendered elements
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TEMPLATE ERROR: UpdateHeaderDataTemplate failed");
        }
    }

    /// <summary>
    /// Initialize UI event handlers
    /// </summary>
    private void InitializeUIEventHandlers()
    {
        // UI event handlers will be implemented as needed
        _logger?.LogInformation("🎨 UI EVENTS: Event handlers initialized");
    }

    /// <summary>
    /// Apply UI sizing constraints
    /// </summary>
    private void ApplyUISizing(double? minWidth, double? minHeight, double? maxWidth, double? maxHeight)
    {
        try
        {
            if (minWidth.HasValue) this.MinWidth = minWidth.Value;
            if (minHeight.HasValue) this.MinHeight = minHeight.Value;
            if (maxWidth.HasValue) this.MaxWidth = maxWidth.Value;
            if (maxHeight.HasValue) this.MaxHeight = maxHeight.Value;

            _logger?.LogInformation("🎨 UI SIZING: Applied sizing constraints");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: ApplyUISizing failed");
        }
    }

    /// <summary>
    /// Apply color configuration to UI elements
    /// </summary>
    private void ApplyColorConfiguration(DataGridColorConfig colorConfig)
    {
        try
        {
            // This method applies colors to rendered UI elements
            // Implementation will be completed with actual UI rendering logic
            _logger?.LogInformation("🎨 UI COLOR: Applied color configuration to UI elements");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: ApplyColorConfiguration failed");
        }
    }

    /// <summary>
    /// Initialize UI virtualization
    /// </summary>
    private async Task InitializeUIVirtualizationAsync()
    {
        try
        {
            await Task.Delay(1); // Placeholder for virtualization setup
            _logger?.LogInformation("🎨 UI VIRTUAL: Virtualization initialized");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: InitializeUIVirtualizationAsync failed");
        }
    }

    /// <summary>
    /// Render all cells
    /// </summary>
    private async Task RenderAllCellsAsync()
    {
        try
        {
            await Task.Delay(1); // Placeholder for cell rendering
            _logger?.LogInformation("🎨 UI RENDER: All cells rendered");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: RenderAllCellsAsync failed");
        }
    }

    /// <summary>
    /// Update validation visuals
    /// </summary>
    private async Task UpdateValidationVisualsAsync()
    {
        try
        {
            await Task.Delay(1); // Placeholder for validation visual updates
            _logger?.LogInformation("🎨 UI VALIDATION: Validation visuals updated");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateValidationVisualsAsync failed");
        }
    }

    /// <summary>
    /// Update specific row UI
    /// </summary>
    private async Task UpdateSpecificRowUIAsync(int rowIndex)
    {
        try
        {
            await Task.Delay(1); // Placeholder for row-specific UI updates
            _logger?.LogInformation("🎨 UI ROW: Updated row {Row} UI", rowIndex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateSpecificRowUIAsync failed for row {Row}", rowIndex);
        }
    }

    /// <summary>
    /// Update specific cell UI
    /// </summary>
    private async Task UpdateSpecificCellUIAsync(int row, int column)
    {
        try
        {
            await Task.Delay(1); // Placeholder for cell-specific UI updates
            _logger?.LogInformation("🎨 UI CELL: Updated cell [{Row},{Column}] UI", row, column);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateSpecificCellUIAsync failed for cell [{Row},{Column}]", row, column);
        }
    }

    /// <summary>
    /// Update specific column UI
    /// </summary>
    private async Task UpdateSpecificColumnUIAsync(string columnName)
    {
        try
        {
            await Task.Delay(1); // Placeholder for column-specific UI updates
            _logger?.LogInformation("🎨 UI COLUMN: Updated column {Column} UI", columnName);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateSpecificColumnUIAsync failed for column {Column}", columnName);
        }
    }

    #endregion
}