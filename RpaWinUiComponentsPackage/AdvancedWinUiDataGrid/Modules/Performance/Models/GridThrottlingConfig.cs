namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;

/// <summary>
/// Performance throttling konfigurácia pre DataGrid
/// </summary>
public class GridThrottlingConfig
{
    /// <summary>
    /// Interval pre UI updates (ms) - ako často sa môže UI refreshovať
    /// </summary>
    public int UIUpdateIntervalMs { get; set; } = 16; // ~60 FPS

    /// <summary>
    /// Interval pre validation updates (ms) - ako často sa spúšťa real-time validation
    /// </summary>
    public int ValidationUpdateIntervalMs { get; set; } = 100;

    /// <summary>
    /// Interval pre search updates (ms) - debouncing pre search input
    /// </summary>
    public int SearchUpdateIntervalMs { get; set; } = 300;

    /// <summary>
    /// Batch size pre bulk operations (koľko riadkov spracovať naraz)
    /// </summary>
    public int BulkOperationBatchSize { get; set; } = 1000;

    /// <summary>
    /// Maximálny počet visible rows pre virtualization
    /// </summary>
    public int VirtualizationBufferSize { get; set; } = 50;

    /// <summary>
    /// Timeout pre UI operácie (ms)
    /// </summary>
    public int UIOperationTimeoutMs { get; set; } = 5000; // 5 sekúnd

    /// <summary>
    /// Timeout pre data operácie (ms)
    /// </summary>
    public int DataOperationTimeoutMs { get; set; } = 60000; // 1 minúta

    /// <summary>
    /// Timeout pre import/export operácie (ms)
    /// </summary>
    public int ImportExportTimeoutMs { get; set; } = 300000; // 5 minút

    /// <summary>
    /// Maximálny počet riadkov pre real-time validation (nad týmto počtom sa používa batch validation)
    /// </summary>
    public int MaxRowsForRealtimeValidation { get; set; } = 1000;

    /// <summary>
    /// Interval pre memory cleanup (ms)
    /// </summary>
    public int MemoryCleanupIntervalMs { get; set; } = 30000; // 30 sekúnd

    /// <summary>
    /// Interval pre cache cleanup (ms)
    /// </summary>
    public int CacheCleanupIntervalMs { get; set; } = 60000; // 1 minúta

    /// <summary>
    /// Maximálny počet items v search history
    /// </summary>
    public int MaxSearchHistoryItems { get; set; } = 100;

    /// <summary>
    /// Zapnúť aggressive memory management
    /// </summary>
    public bool EnableAggressiveMemoryManagement { get; set; } = false;

    /// <summary>
    /// Zapnúť multi-level caching
    /// </summary>
    public bool EnableMultiLevelCaching { get; set; } = true;

    /// <summary>
    /// Zapnúť background processing pre non-critical operations
    /// </summary>
    public bool EnableBackgroundProcessing { get; set; } = true;

    /// <summary>
    /// Default konfigurácia (balanced performance)
    /// </summary>
    public static GridThrottlingConfig Default => new();

    /// <summary>
    /// High performance konfigurácia (pre malé datasety)
    /// </summary>
    public static GridThrottlingConfig HighPerformance => new()
    {
        UIUpdateIntervalMs = 8, // ~120 FPS
        ValidationUpdateIntervalMs = 50,
        SearchUpdateIntervalMs = 150,
        BulkOperationBatchSize = 500,
        VirtualizationBufferSize = 100,
        MaxRowsForRealtimeValidation = 2000,
        EnableAggressiveMemoryManagement = false,
        EnableBackgroundProcessing = true
    };

    /// <summary>
    /// Battery saver konfigurácia (pre veľké datasety alebo slabé zariadenia)
    /// </summary>
    public static GridThrottlingConfig BatterySaver => new()
    {
        UIUpdateIntervalMs = 33, // ~30 FPS
        ValidationUpdateIntervalMs = 500,
        SearchUpdateIntervalMs = 500,
        BulkOperationBatchSize = 2000,
        VirtualizationBufferSize = 30,
        MaxRowsForRealtimeValidation = 500,
        MemoryCleanupIntervalMs = 15000, // 15 sekúnd
        EnableAggressiveMemoryManagement = true,
        EnableBackgroundProcessing = true
    };

    /// <summary>
    /// Large dataset konfigurácia (pre veľké datasety)
    /// </summary>
    public static GridThrottlingConfig LargeDataset => new()
    {
        UIUpdateIntervalMs = 20, // ~50 FPS
        ValidationUpdateIntervalMs = 200,
        SearchUpdateIntervalMs = 400,
        BulkOperationBatchSize = 5000,
        VirtualizationBufferSize = 40,
        MaxRowsForRealtimeValidation = 0, // Disable real-time validation
        DataOperationTimeoutMs = 120000, // 2 minúty
        ImportExportTimeoutMs = 600000, // 10 minút
        EnableAggressiveMemoryManagement = true,
        EnableMultiLevelCaching = true,
        EnableBackgroundProcessing = true
    };

    /// <summary>
    /// Validuje konfiguráciu
    /// </summary>
    public bool IsValid(out string? errorMessage)
    {
        if (UIUpdateIntervalMs <= 0)
        {
            errorMessage = "UIUpdateIntervalMs must be > 0";
            return false;
        }

        if (ValidationUpdateIntervalMs <= 0)
        {
            errorMessage = "ValidationUpdateIntervalMs must be > 0";
            return false;
        }

        if (SearchUpdateIntervalMs <= 0)
        {
            errorMessage = "SearchUpdateIntervalMs must be > 0";
            return false;
        }

        if (BulkOperationBatchSize <= 0)
        {
            errorMessage = "BulkOperationBatchSize must be > 0";
            return false;
        }

        if (VirtualizationBufferSize <= 0)
        {
            errorMessage = "VirtualizationBufferSize must be > 0";
            return false;
        }

        if (MaxRowsForRealtimeValidation < 0)
        {
            errorMessage = "MaxRowsForRealtimeValidation must be >= 0";
            return false;
        }

        errorMessage = null;
        return true;
    }

    /// <summary>
    /// Klonuje throttling config
    /// </summary>
    public GridThrottlingConfig Clone()
    {
        return new GridThrottlingConfig
        {
            UIUpdateIntervalMs = UIUpdateIntervalMs,
            ValidationUpdateIntervalMs = ValidationUpdateIntervalMs,
            SearchUpdateIntervalMs = SearchUpdateIntervalMs,
            BulkOperationBatchSize = BulkOperationBatchSize,
            VirtualizationBufferSize = VirtualizationBufferSize,
            UIOperationTimeoutMs = UIOperationTimeoutMs,
            DataOperationTimeoutMs = DataOperationTimeoutMs,
            ImportExportTimeoutMs = ImportExportTimeoutMs,
            MaxRowsForRealtimeValidation = MaxRowsForRealtimeValidation,
            MemoryCleanupIntervalMs = MemoryCleanupIntervalMs,
            CacheCleanupIntervalMs = CacheCleanupIntervalMs,
            MaxSearchHistoryItems = MaxSearchHistoryItems,
            EnableAggressiveMemoryManagement = EnableAggressiveMemoryManagement,
            EnableMultiLevelCaching = EnableMultiLevelCaching,
            EnableBackgroundProcessing = EnableBackgroundProcessing
        };
    }
}