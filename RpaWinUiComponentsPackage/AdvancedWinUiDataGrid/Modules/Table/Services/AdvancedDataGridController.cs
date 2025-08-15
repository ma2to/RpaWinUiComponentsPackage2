using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml.Controls;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Services;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Validation.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Services;

/// <summary>
/// AdvancedDataGrid Controller - business logic controller pre UI wrapper
/// Oddeltá business logika od UI komponenty, čistá separation of concerns
/// Controller pattern pre AdvancedDataGrid functionality
/// </summary>
public class AdvancedDataGridController
{
    #region Private Fields

    /// <summary>
    /// Headless core - všetky business operácie
    /// </summary>
    private readonly DynamicTableCore _tableCore;

    /// <summary>
    /// Smart column name resolver pre duplicate handling
    /// </summary>
    private SmartColumnNameResolver? _columnNameResolver;

    /// <summary>
    /// Unlimited row height manager pre content overflow handling
    /// </summary>
    private UnlimitedRowHeightManager? _rowHeightManager;

    /// <summary>
    /// Zebra row color manager pre runtime color theming
    /// </summary>
    private ZebraRowColorManager? _zebraColorManager;

    /// <summary>
    /// Logger (nullable - funguje bez loggera)
    /// </summary>
    private Microsoft.Extensions.Logging.ILogger? _logger;

    /// <summary>
    /// Je controller inicializovaný
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

    #region Constructor

    public AdvancedDataGridController()
    {
        _tableCore = new DynamicTableCore();
    }

    #endregion

    #region Properties

    /// <summary>
    /// Je controller inicializovaný
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

    /// <summary>
    /// Current color configuration
    /// </summary>
    public DataGridColorConfig ColorConfig => _colorConfig;

    /// <summary>
    /// Current row height
    /// </summary>
    public double CurrentRowHeight => _rowHeightManager?.CurrentUnifiedRowHeight ?? 32.0;

    /// <summary>
    /// Je zebra coloring enabled
    /// </summary>
    public bool IsZebraColoringEnabled => _zebraColorManager?.IsZebraColoringEnabled ?? true;

    /// <summary>
    /// Zebra color manager (pre advanced scenarios)
    /// </summary>
    public ZebraRowColorManager? ZebraColorManager => _zebraColorManager;

    #endregion

    #region Initialization

    /// <summary>
    /// Inicializuje controller s column definitions
    /// HLAVNÝ ENTRY POINT pre všetko použitie - business logic level
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
        int searchHistoryItems = 0)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            _logger = logger;
            _throttlingConfig = throttlingConfig ?? GridThrottlingConfig.Default;
            _colorConfig = colorConfig ?? DataGridColorConfig.Default;

            // ROZSIAHLE LOGOVANIE - method entry s parametrami
            _logger?.LogMethodEntry(nameof(InitializeAsync), 
                columns?.Count, validationConfig?.IsValidationEnabled, emptyRowsCount, 
                enableBatchValidation, enableSort, enableSearch, enableFilter);
            
            // Log data details
            _logger?.LogDataDetails("COLUMNS_INPUT", columns, columns?.Count);

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

            _isInitialized = true;

            // ROZSIAHLE LOGOVANIE - method exit s časovaním
            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(InitializeAsync), stopwatch.Elapsed, "SUCCESS");
            _logger?.LogPerformance("INITIALIZATION", stopwatch.Elapsed, columns?.Count);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(InitializeAsync), stopwatch.Elapsed, "FAILED");
            _logger?.Error(ex, "🚨 CONTROLLER INIT ERROR: AdvancedDataGrid controller initialization failed after {Time}ms", stopwatch.ElapsedMilliseconds);
            throw;
        }
    }

    #endregion

    #region Data Operations - Passthrough to Core

    /// <summary>
    /// Import z Dictionary list - HEADLESS operation (NO automatic UI refresh)
    /// </summary>
    public async Task ImportFromDictionaryAsync(
        List<Dictionary<string, object?>> data,
        Dictionary<int, bool>? checkboxStates = null,
        int? startRow = null,
        bool insertMode = false,
        TimeSpan? timeout = null,
        IProgress<ValidationProgress>? validationProgress = null)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            if (!_isInitialized)
                throw new InvalidOperationException("Controller must be initialized first");

            // ROZSIAHLE LOGOVANIE - method entry s parametrami
            _logger?.LogMethodEntry(nameof(ImportFromDictionaryAsync), 
                data?.Count, checkboxStates?.Count, startRow, insertMode, timeout?.TotalSeconds);
            
            // Log import data details
            _logger?.LogDataDetails("IMPORT_DATA", data, data?.Count);
            if (checkboxStates?.Any() == true)
            {
                _logger?.LogDataDetails("CHECKBOX_STATES", checkboxStates, checkboxStates.Count);
            }

            // Direct passthrough to headless core
            await _tableCore.ImportFromDictionaryAsync(data, checkboxStates, startRow, insertMode, timeout, validationProgress);

            // Auto-recalculate row height after import (if needed)
            if (_rowHeightManager?.IsEnabled == true)
            {
                var heightStopwatch = System.Diagnostics.Stopwatch.StartNew();
                await _rowHeightManager.RecalculateUnifiedRowHeightAsync(data);
                heightStopwatch.Stop();
                _logger?.LogPerformance("ROW_HEIGHT_RECALC", heightStopwatch.Elapsed, data?.Count);
            }

            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(ImportFromDictionaryAsync), stopwatch.Elapsed, "SUCCESS_HEADLESS");
            _logger?.LogPerformance("DATA_IMPORT", stopwatch.Elapsed, data?.Count);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(ImportFromDictionaryAsync), stopwatch.Elapsed, "FAILED");
            _logger?.Error(ex, "🚨 CONTROLLER IMPORT ERROR: ImportFromDictionaryAsync failed after {Time}ms with {Count} rows", 
                          stopwatch.ElapsedMilliseconds, data?.Count ?? 0);
            throw;
        }
    }

    /// <summary>
    /// Export do Dictionary list - HEADLESS operation
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ExportToDictionaryAsync(
        bool includeValidAlerts = false,
        bool removeAfter = false,
        TimeSpan? timeout = null,
        IProgress<ExportProgress>? exportProgress = null)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("Controller must be initialized first");

        return await _tableCore.ExportToDictionaryAsync(includeValidAlerts, removeAfter, timeout, exportProgress);
    }

    /// <summary>
    /// Nastaví hodnotu bunky - HEADLESS operation
    /// </summary>
    public async Task SetCellValueAsync(int row, int column, object? value)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            if (!_isInitialized)
                throw new InvalidOperationException("Controller must be initialized first");

            // ROZSIAHLE LOGOVANIE - method entry s parametrami
            var valuePreview = value?.ToString()?.Length > 50 ? 
                value.ToString()!.Substring(0, 50) + "..." : value?.ToString();
            _logger?.LogMethodEntry(nameof(SetCellValueAsync), row, column, value?.GetType().Name, valuePreview);

            // Direct passthrough to headless core
            await _tableCore.SetCellValueAsync(row, column, value);

            // Check if row height recalculation is needed for this new value
            if (_rowHeightManager?.IsEnabled == true)
            {
                var columnNames = _tableCore.GetColumnNames();
                if (column >= 0 && column < columnNames.Count)
                {
                    var columnName = columnNames[column];
                    var heightStopwatch = System.Diagnostics.Stopwatch.StartNew();
                    var needsRecalculation = await _rowHeightManager.CheckIfRecalculationNeededAsync(value, columnName);
                    
                    if (needsRecalculation)
                    {
                        var currentData = await _tableCore.ExportToDictionaryAsync();
                        await _rowHeightManager.RecalculateUnifiedRowHeightAsync(currentData);
                        heightStopwatch.Stop();
                        _logger?.LogPerformance("AUTO_HEIGHT_RECALC", heightStopwatch.Elapsed, ActualRowCount);
                    }
                }
            }

            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(SetCellValueAsync), stopwatch.Elapsed, "SUCCESS_HEADLESS");
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(SetCellValueAsync), stopwatch.Elapsed, "FAILED");
            _logger?.Error(ex, "🚨 CONTROLLER CELL ERROR: SetCellValueAsync failed after {Time}ms - Row: {Row}, Column: {Column}", 
                          stopwatch.ElapsedMilliseconds, row, column);
            throw;
        }
    }

    #endregion

    #region Color Configuration

    /// <summary>
    /// Aplikuje nové farby okamžite - SELECTIVE MERGE (len nastavené farby, zvyšok zostane default)
    /// </summary>
    public void ApplyColorConfig(DataGridColorConfig? colorConfig = null)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // ROZSIAHLE LOGOVANIE - method entry
            _logger?.LogMethodEntry(nameof(ApplyColorConfig), 
                colorConfig?.GetType().Name ?? "null",
                colorConfig?.CellBorderColor?.ToString(),
                colorConfig?.SelectionBackgroundColor?.ToString(),
                colorConfig?.ValidationErrorBorderColor?.ToString());

            // Default approach: ak null, zachovaj aktuálne (default) farby
            if (colorConfig == null) 
            {
                _logger?.LogMethodExit(nameof(ApplyColorConfig), stopwatch.Elapsed, "NO_COLORS_KEEPING_DEFAULTS");
                return;
            }

            // Count how many colors are being set by application
            var colorPropertiesCount = typeof(DataGridColorConfig).GetProperties()
                .Count(prop => prop.PropertyType == typeof(Windows.UI.Color?) && 
                              prop.GetValue(colorConfig) != null);

            _logger?.LogDataDetails("COLOR_CONFIG_INPUT", colorConfig, colorPropertiesCount);

            // Selective merge: aplikuj len non-null farby z aplikácie, zvyšok nechaj default
            _colorConfig.MergeWith(colorConfig);
            
            // Apply to zebra color manager
            _zebraColorManager?.ApplyColorConfiguration(_colorConfig);

            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(ApplyColorConfig), stopwatch.Elapsed, 
                $"SELECTIVE_MERGE_{colorPropertiesCount}_COLORS");
            _logger?.LogPerformance("COLOR_CONFIG_APPLY", stopwatch.Elapsed, colorPropertiesCount);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(ApplyColorConfig), stopwatch.Elapsed, "FAILED");
            _logger?.Error(ex, "🚨 CONTROLLER COLOR ERROR: ApplyColorConfig failed after {Time}ms", stopwatch.ElapsedMilliseconds);
            throw;
        }
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

    #endregion

    #region Validation

    /// <summary>
    /// Validuje VŠETKY riadky v dataset - HEADLESS operation
    /// </summary>
    public async Task<bool> AreAllNonEmptyRowsValidAsync()
    {
        if (!_isInitialized)
            throw new InvalidOperationException("Controller must be initialized first");

        return await _tableCore.AreAllNonEmptyRowsValidAsync();
    }

    /// <summary>
    /// Batch validation všetkých riadkov - HEADLESS operation
    /// </summary>
    public async Task<BatchValidationResult?> ValidateAllRowsBatchAsync(CancellationToken cancellationToken = default)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("Controller must be initialized first");

        var result = await _tableCore.ValidateAllRowsBatchAsync(cancellationToken);
        
        _logger?.Info("📊 CONTROLLER VALIDATION: ValidateAllRowsBatchAsync completed (HEADLESS)");
        
        return result;
    }

    #endregion

    #region Column Management

    /// <summary>
    /// Získa všetky resolved column names (po duplicate resolution)
    /// </summary>
    public List<string> GetResolvedColumnNames()
    {
        if (!_isInitialized)
            throw new InvalidOperationException("Controller must be initialized first");

        try
        {
            return _tableCore.GetColumnNames();
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CONTROLLER ERROR: GetResolvedColumnNames failed");
            throw;
        }
    }

    #endregion

    #region Row Height Management

    /// <summary>
    /// Recalculate row height pre current dataset
    /// </summary>
    public async Task RecalculateRowHeightAsync()
    {
        if (!_isInitialized || _rowHeightManager == null) return;

        try
        {
            _logger?.Info("📐 CONTROLLER: Manual row height recalculation started");

            // Get current data from table core
            var currentData = await _tableCore.ExportToDictionaryAsync();
            
            // Recalculate unified height
            var newHeight = await _rowHeightManager.RecalculateUnifiedRowHeightAsync(currentData);
            
            _logger?.Info("✅ CONTROLLER: Row height recalculated - New height: {Height}px", Math.Ceiling(newHeight));
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CONTROLLER ERROR: RecalculateRowHeightAsync failed");
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
            _logger?.Info("📐 CONTROLLER: Base row height set to {Height}px", baseHeight);
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
            _logger?.Info("📐 CONTROLLER: Unlimited row height {Status}", enabled ? "ENABLED" : "DISABLED");
        }
    }

    #endregion
}