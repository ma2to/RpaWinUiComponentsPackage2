using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Services.Core;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Configuration;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Validation;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Controls;

/// <summary>
/// Advanced WinUI DataGrid - UI Wrapper okolo DynamicTableCore
/// Partial class architecture implementovaná od začiatku
/// Implementuje Dual Usage Modes - UI Application Mode
/// </summary>
public sealed partial class AdvancedDataGrid : UserControl
{
    #region Private Fields

    /// <summary>
    /// Headless core - všetky business operácie
    /// </summary>
    private DynamicTableCore _tableCore;

    /// <summary>
    /// Smart column name resolver pre duplicate handling
    /// </summary>
    private SmartColumnNameResolver _columnNameResolver;

    /// <summary>
    /// Unlimited row height manager pre content overflow handling
    /// </summary>
    private UnlimitedRowHeightManager _rowHeightManager;

    /// <summary>
    /// Zebra row color manager pre runtime color theming
    /// </summary>
    private ZebraRowColorManager _zebraColorManager;

    /// <summary>
    /// Logger (nullable - funguje bez loggera)
    /// </summary>
    private Microsoft.Extensions.Logging.ILogger? _logger;

    /// <summary>
    /// Je komponent inicializovaný
    /// </summary>
    private bool _isInitialized = false;

    /// <summary>
    /// Color configuration
    /// </summary>
    private DataGridColorConfig _colorConfig = DataGridColorConfig.Default;

    /// <summary>
    /// Throttling configuration
    /// </summary>
    private GridThrottlingConfig _throttlingConfig = GridThrottlingConfig.Default;

    #endregion

    #region Properties

    /// <summary>
    /// Je komponent inicializovaný
    /// </summary>
    public bool IsInitialized => _isInitialized;

    /// <summary>
    /// Headless table core (for advanced scenarios)
    /// </summary>
    public DynamicTableCore TableCore => _tableCore;

    /// <summary>
    /// Aktuálny počet riadkov
    /// </summary>
    public int ActualRowCount => _tableCore?.ActualRowCount ?? 0;

    /// <summary>
    /// Počet stĺpcov
    /// </summary>
    public int ColumnCount => _tableCore?.ColumnCount ?? 0;

    /// <summary>
    /// Má dataset nejaké dáta
    /// </summary>
    public bool HasData => _tableCore?.HasData ?? false;

    #endregion

    #region Constructor

    public AdvancedDataGrid()
    {
        this.InitializeComponent();
        _tableCore = new DynamicTableCore();
        
        // Initialize services (logger will be set during InitializeAsync)
        _columnNameResolver = new SmartColumnNameResolver();
        _rowHeightManager = new UnlimitedRowHeightManager();
        _zebraColorManager = new ZebraRowColorManager();
        
        // Initialize UI event handlers
        InitializeUIEventHandlers();
    }

    #endregion

    #region Public API - Initialization

    /// <summary>
    /// Inicializuje DataGrid s column definitions
    /// HLAVNÝ ENTRY POINT pre všetko použitie
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
            _throttlingConfig = throttlingConfig ?? GridThrottlingConfig.Default;
            _colorConfig = colorConfig ?? DataGridColorConfig.Default;

            _logger?.Info("🚀 OPERATION START: AdvancedDataGrid.InitializeAsync - Columns: {Count}, Rules: {Rules}", 
                                   columns.Count, validationConfig?.IsValidationEnabled ?? false);

            // Phase 1: Resolve duplicate column names
            _columnNameResolver = new SmartColumnNameResolver(_logger);
            var resolvedColumns = _columnNameResolver.ResolveDuplicateNames(columns);
            
            _logger?.Info("🔧 DUPLICATE RESOLUTION: {OriginalCount} → {ResolvedCount} columns processed", 
                columns.Count, resolvedColumns.Count);

            // Phase 2: Initialize row height manager
            _rowHeightManager = new UnlimitedRowHeightManager(_logger);
            var fontInfo = new FontInfo(); // Default font info, can be configured later
            _rowHeightManager.Initialize(resolvedColumns, fontInfo);

            // Phase 3: Initialize zebra color manager
            _zebraColorManager = new ZebraRowColorManager(_logger);
            _zebraColorManager.ApplyColorConfiguration(_colorConfig);

            // Phase 4: Initialize headless core with resolved columns
            await _tableCore.InitializeAsync(resolvedColumns, validationConfig, _throttlingConfig, emptyRowsCount, _logger);

            // Apply UI sizing
            ApplyUISizing(minWidth, minHeight, maxWidth, maxHeight);

            // Apply color configuration
            ApplyColorConfiguration(_colorConfig);

            // Initialize UI virtualization
            await InitializeUIVirtualizationAsync();

            _isInitialized = true;

            _logger?.Info("✅ OPERATION SUCCESS: AdvancedDataGrid initialized - UI + Core ready");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 INIT ERROR: AdvancedDataGrid initialization failed");
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
            _logger?.Info("🎨 UI UPDATE: Full UI refresh started");
            
            await ShowLoadingAsync(true);
            await RenderAllCellsAsync();
            await ShowLoadingAsync(false);

            _logger?.Info("✅ UI UPDATE: Full UI refresh completed");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: RefreshUIAsync failed");
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
            _logger?.Info("🎨 UI UPDATE: Validation UI update started");

            await UpdateValidationVisualsAsync();

            _logger?.Info("✅ UI UPDATE: Validation UI update completed");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateValidationUIAsync failed");
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
            _logger?.Info("🎨 UI UPDATE: Row UI update - Row: {Row}", rowIndex);

            await UpdateSpecificRowUIAsync(rowIndex);

            _logger?.Info("✅ UI UPDATE: Row UI update completed - Row: {Row}", rowIndex);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateRowUIAsync failed - Row: {Row}", rowIndex);
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
            _logger?.Info("🎨 UI UPDATE: Cell UI update - Row: {Row}, Column: {Column}", row, column);

            await UpdateSpecificCellUIAsync(row, column);

            _logger?.Info("✅ UI UPDATE: Cell UI update completed - Row: {Row}, Column: {Column}", row, column);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateCellUIAsync failed - Row: {Row}, Column: {Column}", row, column);
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
            _logger?.Info("🎨 UI UPDATE: Column UI update - Column: {Column}", columnName);

            await UpdateSpecificColumnUIAsync(columnName);

            _logger?.Info("✅ UI UPDATE: Column UI update completed - Column: {Column}", columnName);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateColumnUIAsync failed - Column: {Column}", columnName);
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
            _logger?.Info("🎨 UI UPDATE: Layout invalidation");

            // Update layout with current row height (if available)
            if (DataRepeater != null && _rowHeightManager != null)
            {
                var currentRowHeight = _rowHeightManager.CurrentUnifiedRowHeight;
                var layout = new UniformGridLayout
                {
                    Orientation = Orientation.Vertical,
                    MinItemWidth = 120,
                    MinItemHeight = currentRowHeight,
                    ItemsStretch = UniformGridLayoutItemsStretch.Fill
                };
                DataRepeater.Layout = layout;
                
                _logger?.Info("🎨 UI LAYOUT: Updated with unified row height: {Height}px", Math.Ceiling(currentRowHeight));
            }

            // Force ItemsRepeater to recalculate layout
            DataRepeater?.InvalidateMeasure();
            HeaderRepeater?.InvalidateMeasure();

            _logger?.Info("✅ UI UPDATE: Layout invalidation completed");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: InvalidateLayout failed");
        }
    }

    #endregion

    #region Passthrough API to DynamicTableCore (Headless Operations)

    /// <summary>
    /// Import z Dictionary list - HEADLESS operation (NO automatic UI refresh)
    /// Pre UI aplikácie: po import volaj RefreshUIAsync() manuálne
    /// </summary>
    public async Task ImportFromDictionaryAsync(
        List<Dictionary<string, object?>> data,
        Dictionary<int, bool>? checkboxStates = null,
        int? startRow = null,
        bool insertMode = false,
        TimeSpan? timeout = null,
        IProgress<ValidationProgress>? validationProgress = null)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // Direct passthrough to headless core (NO UI update)
        await _tableCore.ImportFromDictionaryAsync(data, checkboxStates, startRow, insertMode, timeout, validationProgress);

        // Auto-recalculate row height after import (if needed)
        if (_rowHeightManager?.IsEnabled == true)
        {
            await _rowHeightManager.RecalculateUnifiedRowHeightAsync(data);
            _logger?.Info("📐 AUTO OPERATION: Row height recalculated after data import");
        }

        _logger?.Info("📊 API OPERATION: ImportFromDictionaryAsync completed (HEADLESS) - call RefreshUIAsync() for UI update");
    }

    /// <summary>
    /// Export do Dictionary list - HEADLESS operation
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ExportToDictionaryAsync(
        bool includeValidAlerts = false,
        TimeSpan? timeout = null,
        IProgress<ExportProgress>? exportProgress = null)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // Direct passthrough to headless core
        return await _tableCore.ExportToDictionaryAsync(includeValidAlerts, timeout, exportProgress);
    }

    /// <summary>
    /// Validuje VŠETKY riadky v dataset - HEADLESS operation (NO automatic UI refresh)
    /// Pre UI aplikácie: po validation volaj UpdateValidationUIAsync() manuálne
    /// </summary>
    public async Task<bool> AreAllNonEmptyRowsValidAsync()
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // Direct passthrough to headless core (NO UI update)
        return await _tableCore.AreAllNonEmptyRowsValidAsync();
    }

    /// <summary>
    /// Batch validation všetkých riadkov - HEADLESS operation (NO automatic UI refresh)
    /// Pre UI aplikácie: po validation volaj UpdateValidationUIAsync() manuálne
    /// </summary>
    public async Task<BatchValidationResult?> ValidateAllRowsBatchAsync(CancellationToken cancellationToken = default)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // Direct passthrough to headless core (NO UI update)
        var result = await _tableCore.ValidateAllRowsBatchAsync(cancellationToken);
        
        _logger?.Info("📊 API OPERATION: ValidateAllRowsBatchAsync completed (HEADLESS) - call UpdateValidationUIAsync() for visual indicators");
        
        return result;
    }

    /// <summary>
    /// Nastaví hodnotu bunky - HEADLESS operation (NO automatic UI refresh)
    /// Pre UI aplikácie: po zmene volaj UpdateCellUIAsync() manuálne
    /// </summary>
    public async Task SetCellValueAsync(int row, int column, object? value)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // Direct passthrough to headless core (NO UI update)
        await _tableCore.SetCellValueAsync(row, column, value);

        // Check if row height recalculation is needed for this new value
        if (_rowHeightManager?.IsEnabled == true)
        {
            var columnNames = _tableCore.GetColumnNames();
            if (column >= 0 && column < columnNames.Count)
            {
                var columnName = columnNames[column];
                var needsRecalculation = await _rowHeightManager.CheckIfRecalculationNeededAsync(value, columnName);
                
                if (needsRecalculation)
                {
                    var currentData = await _tableCore.ExportToDictionaryAsync();
                    await _rowHeightManager.RecalculateUnifiedRowHeightAsync(currentData);
                    _logger?.Info("📐 AUTO OPERATION: Row height recalculated after cell value change");
                }
            }
        }

        _logger?.Info("📊 API OPERATION: SetCellValueAsync completed (HEADLESS) - call UpdateCellUIAsync({Row}, {Column}) for UI update", row, column);
    }

    #endregion

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

    #region Color Configuration API

    /// <summary>
    /// Aplikuje nové farby okamžite (prepíše initialization farby)
    /// </summary>
    public void ApplyColorConfig(DataGridColorConfig colorConfig)
    {
        _colorConfig = colorConfig;
        
        // Apply to UI infrastructure
        ApplyColorConfiguration(_colorConfig);
        
        // Apply to zebra color manager
        _zebraColorManager?.ApplyColorConfiguration(_colorConfig);

        _logger?.Info("🎨 COLOR CONFIG: Applied new color configuration (UI + Zebra colors)");
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
        if (_zebraColorManager != null)
        {
            _zebraColorManager.IsZebraColoringEnabled = enabled;
            _logger?.Info("🎨 COLOR CONFIG: Zebra coloring {Status}", enabled ? "ENABLED" : "DISABLED");
        }
    }

    /// <summary>
    /// Je zebra coloring enabled
    /// </summary>
    public bool IsZebraColoringEnabled => _zebraColorManager?.IsZebraColoringEnabled ?? true;

    /// <summary>
    /// Získa zebra color manager (pre advanced scenarios)
    /// </summary>
    public ZebraRowColorManager ZebraColorManager => _zebraColorManager;

    #endregion

    #region Column Names API

    /// <summary>
    /// Získa všetky resolved column names (po duplicate resolution)
    /// Používa sa pre business logiku namiesto display names
    /// </summary>
    public List<string> GetResolvedColumnNames()
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        try
        {
            return _tableCore.GetColumnNames();
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 API ERROR: GetResolvedColumnNames failed");
            throw;
        }
    }

    /// <summary>
    /// Vráti info o tom či column name bol premenovaný during duplicate resolution
    /// </summary>
    public Dictionary<string, string> GetColumnNameMappings()
    {
        if (!_isInitialized)
            throw new InvalidOperationException("DataGrid must be initialized first");

        // This would require storing the original->resolved mapping in SmartColumnNameResolver
        // For now, return empty mapping - can be enhanced later
        return new Dictionary<string, string>();
    }

    #endregion

    #region Row Height Management API

    /// <summary>
    /// Aktuálna unified row height (všetky riadky majú rovnakú výšku)
    /// </summary>
    public double CurrentRowHeight => _rowHeightManager?.CurrentUnifiedRowHeight ?? 32.0;

    /// <summary>
    /// Recalculate row height pre current dataset
    /// Volá sa automaticky pri import, ale môže sa volať manuálne
    /// </summary>
    public async Task RecalculateRowHeightAsync()
    {
        if (!_isInitialized || _rowHeightManager == null) return;

        try
        {
            _logger?.Info("📐 API OPERATION: Manual row height recalculation started");

            // Get current data from table core
            var currentData = await _tableCore.ExportToDictionaryAsync();
            
            // Recalculate unified height
            var newHeight = await _rowHeightManager.RecalculateUnifiedRowHeightAsync(currentData);
            
            // Update UI layout
            InvalidateLayout();
            
            _logger?.Info("✅ API OPERATION: Row height recalculated - New height: {Height}px", Math.Ceiling(newHeight));
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 API ERROR: RecalculateRowHeightAsync failed");
        }
    }

    /// <summary>
    /// Set base row height (default pre prázdne bunky)
    /// </summary>
    public void SetBaseRowHeight(double baseHeight)
    {
        if (_rowHeightManager != null)
        {
            _rowHeightManager.BaseRowHeight = baseHeight;
            _logger?.Info("📐 API OPERATION: Base row height set to {Height}px", baseHeight);
        }
    }

    /// <summary>
    /// Enable/disable unlimited row height system
    /// </summary>
    public void SetUnlimitedRowHeightEnabled(bool enabled)
    {
        if (_rowHeightManager != null)
        {
            _rowHeightManager.IsEnabled = enabled;
            _logger?.Info("📐 API OPERATION: Unlimited row height {Status}", enabled ? "ENABLED" : "DISABLED");
        }
    }

    #endregion
}