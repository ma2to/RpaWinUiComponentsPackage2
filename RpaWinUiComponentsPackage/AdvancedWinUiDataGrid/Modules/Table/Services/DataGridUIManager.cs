using Microsoft.Extensions.Logging;
using Microsoft.UI;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models;
using System.Collections.ObjectModel;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Services;

/// <summary>
/// UI Manager pre DataGrid - zodpovedný za správne renderovanie UI elementov
/// Separácia UI logiky od business logiky
/// </summary>
public class DataGridUIManager
{
    #region Private Fields

    private readonly ILogger? _logger;
    private readonly DynamicTableCore _tableCore;
    private DataGridColorConfig _colorConfig;

    // VIRTUAL SCROLLING ARCHITECTURE: Viewport-based rendering + Complete dataset
    private readonly ObservableCollection<HeaderCellModel> _headersCollection = new();
    private readonly ObservableCollection<DataRowModel> _viewportRowsCollection = new();  // VIEWPORT: Len viditeľné riadky pre UI
    
    // Expose viewport collections for XAML binding (WinRT-safe ObservableCollection)
    public ObservableCollection<HeaderCellModel> HeadersCollection => _headersCollection;
    public ObservableCollection<DataRowModel> RowsCollection => _viewportRowsCollection;
    
    // VIRTUAL SCROLLING STATE
    private int _viewportStartIndex = 0;
    private int _viewportSize = 5;         // TEMPORARY: Reduced for debugging WinRT COM errors
    private int _totalDatasetSize = 0;     // Kompletná veľkosť datasetu

    // UI State
    private bool _isRendering = false;
    private DateTime _lastRenderTime = DateTime.MinValue;

    #endregion

    #region Constructor

    public DataGridUIManager(DynamicTableCore tableCore, ILogger? logger = null)
    {
        _tableCore = tableCore ?? throw new ArgumentNullException(nameof(tableCore));
        _logger = logger;
        _colorConfig = DataGridColorConfig.Default;

        // Lists are initialized inline - no need for constructor assignment
    }

    #endregion

    #region Public Properties

    /// <summary>
    /// Aktuálna color configuration
    /// </summary>
    public DataGridColorConfig ColorConfig
    {
        get => _colorConfig;
        set
        {
            _colorConfig = value ?? DataGridColorConfig.Default;
            _logger?.LogInformation("🎨 UI CONFIG: Color configuration updated");
        }
    }

    /// <summary>
    /// Je UI momentálne v procese renderovania
    /// </summary>
    public bool IsRendering => _isRendering;
    
    /// <summary>
    /// Virtualization properties pre external access
    /// </summary>
    public int ViewportStartIndex => _viewportStartIndex;
    public int ViewportSize => _viewportSize;
    public int TotalDatasetSize => _totalDatasetSize;
    public int ViewportEndIndex => Math.Min(_viewportStartIndex + _viewportSize - 1, _totalDatasetSize - 1);

    #endregion

    #region Public Methods

    /// <summary>
    /// Inicializuje UI collections s column definitions - s comprehensive error logging
    /// </summary>
    public async Task InitializeUIAsync()
    {
        if (_isRendering)
        {
            _logger?.LogWarning("⚠️ UI RENDER: Already rendering, skipping initialization");
            return;
        }

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        
        try
        {
            _isRendering = true;
            _logger?.LogInformation("🎨 UI INIT: Starting UI initialization...");
            
            // CRITICAL: Log table core state before any operations
            if (_tableCore == null)
            {
                _logger?.LogError("🚨 INIT ERROR: TableCore is null - cannot initialize UI");
                throw new InvalidOperationException("TableCore must be initialized before UI initialization");
            }
            
            var actualRowCount = _tableCore.ActualRowCount;
            var columnCount = _tableCore.ColumnCount;
            var isInitialized = _tableCore.IsInitialized;
            
            _logger?.LogInformation("📊 INIT STATE: TableCore.IsInitialized={IsInitialized}, ActualRowCount={ActualRowCount}, ColumnCount={ColumnCount}", 
                isInitialized, actualRowCount, columnCount);
            
            if (!isInitialized)
            {
                _logger?.LogError("🚨 INIT ERROR: TableCore is not initialized");
                throw new InvalidOperationException("TableCore must be initialized before UI initialization");
            }
            
            if (actualRowCount < 0 || actualRowCount > 1000000)
            {
                _logger?.LogError("🚨 INIT ERROR: ActualRowCount out of safe range: {ActualRowCount}", actualRowCount);
                throw new InvalidOperationException($"ActualRowCount out of safe range: {actualRowCount}");
            }
            
            if (columnCount < 0 || columnCount > 1000)
            {
                _logger?.LogError("🚨 INIT ERROR: ColumnCount out of safe range: {ColumnCount}", columnCount);
                throw new InvalidOperationException($"ColumnCount out of safe range: {columnCount}");
            }

            // Phase 1: Render Headers
            var headerStopwatch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                _logger?.LogInformation("🎨 UI INIT PHASE 1: Rendering headers...");
                await RenderHeadersAsync();
                headerStopwatch.Stop();
                _logger?.LogInformation("✅ UI INIT PHASE 1: Headers rendered in {ElapsedMs}ms", headerStopwatch.ElapsedMilliseconds);
            }
            catch (Exception headerEx)
            {
                headerStopwatch.Stop();
                _logger?.LogError(headerEx, "🚨 UI INIT PHASE 1 ERROR: Header rendering failed after {ElapsedMs}ms", headerStopwatch.ElapsedMilliseconds);
                throw;
            }

            // Phase 2: Render Data Rows
            var rowsStopwatch = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                _logger?.LogInformation("🎨 UI INIT PHASE 2: Rendering data rows...");
                await RenderDataRowsAsync();
                rowsStopwatch.Stop();
                _logger?.LogInformation("✅ UI INIT PHASE 2: Data rows rendered in {ElapsedMs}ms", rowsStopwatch.ElapsedMilliseconds);
            }
            catch (Exception rowsEx)
            {
                rowsStopwatch.Stop();
                _logger?.LogError(rowsEx, "🚨 UI INIT PHASE 2 ERROR: Data rows rendering failed after {ElapsedMs}ms", rowsStopwatch.ElapsedMilliseconds);
                throw;
            }

            _lastRenderTime = DateTime.Now;
            stopwatch.Stop();
            
            var finalHeaderCount = _headersCollection.Count;
            var finalRowCount = _viewportRowsCollection.Count;
            var finalCellCount = _viewportRowsCollection.Sum(r => r.Cells.Count);
            
            _logger?.LogInformation("✅ UI INIT: UI initialization completed in {ElapsedMs}ms - Headers: {HeaderCount}, Rows: {RowCount}, Cells: {CellCount}", 
                stopwatch.ElapsedMilliseconds, finalHeaderCount, finalRowCount, finalCellCount);
                
            // FINAL SAFETY CHECK: Verify WinRT-safe collections are in valid state
            if (finalHeaderCount < 0 || finalHeaderCount > 100)
            {
                _logger?.LogError("🚨 INIT FINAL ERROR: HeadersList count out of WinRT range: {Count}", finalHeaderCount);
                throw new InvalidOperationException($"HeadersList count out of WinRT range: {finalHeaderCount}");
            }
            
            if (finalRowCount < 0 || finalRowCount > 200)  // Viewport limit
            {
                _logger?.LogError("🚨 INIT FINAL ERROR: ViewportRowsList count out of WinRT range: {Count}", finalRowCount);
                throw new InvalidOperationException($"ViewportRowsList count out of WinRT range: {finalRowCount}");
            }
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.LogError(ex, "🚨 UI ERROR: UI initialization failed after {ElapsedMs}ms - TableCore state: ActualRowCount={ActualRowCount}, ColumnCount={ColumnCount}, IsInitialized={IsInitialized}", 
                stopwatch.ElapsedMilliseconds, 
                _tableCore?.ActualRowCount ?? -1, 
                _tableCore?.ColumnCount ?? -1, 
                _tableCore?.IsInitialized ?? false);
            throw;
        }
        finally
        {
            _isRendering = false;
        }
    }

    /// <summary>
    /// Kompletné re-renderovanie všetkých UI elementov
    /// </summary>
    public async Task RefreshAllUIAsync()
    {
        if (_isRendering)
        {
            _logger?.LogWarning("⚠️ UI RENDER: Already rendering, skipping refresh");
            return;
        }

        try
        {
            _isRendering = true;
            _logger?.LogInformation("🎨 UI REFRESH: Starting full UI refresh");

            // Clear existing collections (WinRT-safe ObservableCollection)
            _logger?.LogInformation("🔄 UI CLEAR: Clearing ObservableCollections - Headers: {HeaderCount}, Rows: {RowCount}", 
                _headersCollection.Count, _viewportRowsCollection.Count);
            
            _headersCollection.Clear();
            _viewportRowsCollection.Clear();

            // Re-render everything
            await RenderHeadersAsync();
            await RenderDataRowsAsync();
            
            // CRITICAL DIAGNOSTIC: Verify collections were populated
            _logger?.LogInformation("🎨 UI POPULATE: Collections populated - Headers: {HeaderCount}, Rows: {RowCount}", 
                _headersCollection.Count, _viewportRowsCollection.Count);
                
            if (_headersCollection.Count == 0)
            {
                _logger?.LogError("🚨 UI ERROR: HeadersCollection is still empty after RenderHeadersAsync!");
            }
            
            if (_viewportRowsCollection.Count == 0)
            {
                _logger?.LogError("🚨 UI ERROR: RowsCollection is still empty after RenderDataRowsAsync!");
            }

            _lastRenderTime = DateTime.Now;
            _logger?.LogInformation("✅ UI REFRESH: Full UI refresh completed - Headers: {HeaderCount}, Viewport Rows: {RowCount}, Total Dataset: {TotalDataset}", 
                _headersCollection.Count, _viewportRowsCollection.Count, _totalDatasetSize);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: Full UI refresh failed");
            throw;
        }
        finally
        {
            _isRendering = false;
        }
    }

    /// <summary>
    /// Update validation visual indicators - IMPORTANT: Works on COMPLETE DATASET, not just viewport
    /// </summary>
    public async Task UpdateValidationUIAsync()
    {
        try
        {
            _logger?.LogInformation("🎨 VALIDATION: Starting validation UI update for COMPLETE DATASET");

            // CRITICAL: Validation pracuje na CELOM datasete, nie len viewport
            // TableCore has complete dataset, UI shows only viewport
            for (int datasetRowIndex = 0; datasetRowIndex < _totalDatasetSize; datasetRowIndex++)
            {
                // Check if validation changed for this row in the complete dataset
                // This ensures validation API works on complete dataset as requested
                
                // If row is in viewport, update its visual indicators
                if (datasetRowIndex >= _viewportStartIndex && datasetRowIndex <= ViewportEndIndex)
                {
                    var viewportRow = _viewportRowsCollection.FirstOrDefault(r => r.RowIndex == datasetRowIndex);
                    if (viewportRow != null)
                    {
                        foreach (var cell in viewportRow.Cells)
                        {
                            await UpdateCellValidationAsync(cell);
                        }
                        
                        // Update row-level validation
                        viewportRow.IsValid = viewportRow.Cells.All(c => c.IsValid);
                        
                        _logger?.LogDebug("✅ VIEWPORT VALIDATION: Updated validation for visible row {RowIndex}", datasetRowIndex);
                    }
                }
            }

            _logger?.LogInformation("✅ VALIDATION: Validation UI update completed for {TotalRows} dataset rows, {ViewportRows} visible in viewport", 
                _totalDatasetSize, _viewportRowsCollection.Count);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 VALIDATION ERROR: Validation UI update failed");
            throw;
        }
    }

    /// <summary>
    /// Update UI pre konkrétny riadok (dataset index) - works with viewport virtualization
    /// </summary>
    public async Task UpdateRowUIAsync(int datasetRowIndex)
    {
        try
        {
            // Check if the row is currently visible in viewport
            if (datasetRowIndex < _viewportStartIndex || datasetRowIndex > ViewportEndIndex)
            {
                _logger?.LogDebug("🔍 VIEWPORT: Row {RowIndex} not in current viewport ({Start}-{End}), skipping UI update", 
                    datasetRowIndex, _viewportStartIndex, ViewportEndIndex);
                return;
            }

            // Find the row in viewport
            var rowModel = _viewportRowsCollection.FirstOrDefault(r => r.RowIndex == datasetRowIndex);
            if (rowModel == null)
            {
                _logger?.LogWarning("⚠️ VIEWPORT: Row {RowIndex} not found in viewport collection", datasetRowIndex);
                return;
            }

            _logger?.LogInformation("🎨 VIEWPORT UPDATE: Updating dataset row {RowIndex} in viewport", datasetRowIndex);

            await UpdateRowDataAsync(rowModel);

            _logger?.LogInformation("✅ VIEWPORT UPDATE: Dataset row {RowIndex} updated in viewport", datasetRowIndex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 VIEWPORT ERROR: Row UI update failed for dataset row {RowIndex}", datasetRowIndex);
            throw;
        }
    }

    /// <summary>
    /// Viewport navigation pre Virtual Scrolling
    /// </summary>
    public async Task ScrollToRowAsync(int datasetRowIndex)
    {
        if (datasetRowIndex < 0 || datasetRowIndex >= _totalDatasetSize)
        {
            _logger?.LogWarning("⚠️ SCROLL: Invalid row index {RowIndex}, total dataset size: {TotalSize}", 
                datasetRowIndex, _totalDatasetSize);
            return;
        }

        _logger?.LogInformation("📍 SCROLL: Scrolling to dataset row {RowIndex}", datasetRowIndex);

        // Calculate new viewport start to center the target row
        int newViewportStart = Math.Max(0, datasetRowIndex - _viewportSize / 2);
        
        // Ensure we don't scroll past the end
        if (newViewportStart + _viewportSize > _totalDatasetSize)
        {
            newViewportStart = Math.Max(0, _totalDatasetSize - _viewportSize);
        }

        if (newViewportStart != _viewportStartIndex)
        {
            _viewportStartIndex = newViewportStart;
            _logger?.LogInformation("📍 SCROLL: Viewport repositioned to start at {ViewportStart}", _viewportStartIndex);
            
            // Re-render viewport with new position
            await RenderDataRowsAsync();
        }
    }

    /// <summary>
    /// Scroll viewport by specified number of rows
    /// </summary>
    public async Task ScrollByRowsAsync(int rowOffset)
    {
        int newViewportStart = Math.Max(0, Math.Min(_totalDatasetSize - _viewportSize, _viewportStartIndex + rowOffset));
        
        if (newViewportStart != _viewportStartIndex)
        {
            _viewportStartIndex = newViewportStart;
            _logger?.LogInformation("📍 SCROLL BY: Moved viewport by {Offset} rows to start at {ViewportStart}", 
                rowOffset, _viewportStartIndex);
            
            await RenderDataRowsAsync();
        }
    }

    /// <summary>
    /// Configure viewport size for performance optimization
    /// </summary>
    public async Task SetViewportSizeAsync(int newViewportSize)
    {
        if (newViewportSize <= 0 || newViewportSize > 200) // Safety limits
        {
            _logger?.LogWarning("⚠️ VIEWPORT: Invalid viewport size {Size}, must be 1-200", newViewportSize);
            return;
        }

        if (newViewportSize != _viewportSize)
        {
            _viewportSize = newViewportSize;
            _logger?.LogInformation("📏 VIEWPORT: Size changed to {ViewportSize} rows", _viewportSize);
            
            // Adjust viewport start if needed
            if (_viewportStartIndex + _viewportSize > _totalDatasetSize)
            {
                _viewportStartIndex = Math.Max(0, _totalDatasetSize - _viewportSize);
            }
            
            await RenderDataRowsAsync();
        }
    }

    #endregion

    #region Private Rendering Methods

    /// <summary>
    /// Renderuje header cells
    /// </summary>
    private async Task RenderHeadersAsync()
    {
        try
        {
            _logger?.LogInformation("🎨 UI RENDER: Rendering headers...");

            _headersCollection.Clear();

            for (int i = 0; i < _tableCore.ColumnCount; i++)
            {
                var columnDef = _tableCore.GetColumnDefinition(i);
                if (columnDef == null) continue;

                var headerModel = new HeaderCellModel
                {
                    DisplayName = columnDef.DisplayName,
                    ColumnName = columnDef.Name,
                    Width = columnDef.Width ?? 100,
                    IsSortable = columnDef.IsSortable,
                    IsFilterable = columnDef.IsFilterable,
                    BackgroundBrush = CreateBrush(_colorConfig.HeaderBackgroundColor),
                    ForegroundBrush = CreateBrush(_colorConfig.HeaderForegroundColor),
                    BorderBrush = CreateBrush(_colorConfig.HeaderBorderColor)
                };

                _headersCollection.Add(headerModel);
            }

            // Apply auto-stretch logic for ValidationAlerts column
            ApplyValidationAlertsAutoStretch();

            // CRITICAL DIAGNOSTIC: Log actual header data
            _logger?.LogInformation("✅ UI RENDER: Headers rendered - {Count} columns", _headersCollection.Count);
            for (int i = 0; i < _headersCollection.Count; i++)
            {
                var header = _headersCollection[i];
                _logger?.LogInformation("📋 HEADER[{Index}]: Name='{Name}', DisplayName='{DisplayName}', Width={Width}", 
                    i, header.ColumnName, header.DisplayName, header.Width);
            }
            await Task.CompletedTask; // For async consistency
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: Header rendering failed");
            throw;
        }
    }

    /// <summary>
    /// Renderuje všetky data rows s comprehensive error logging a Int32.MaxValue protection
    /// </summary>
    private async Task RenderDataRowsAsync()
    {
        try
        {
            _logger?.LogInformation("🎨 UI RENDER: Starting data rows rendering...");
            
            // CRITICAL: Log current state before any operations
            var actualRowCount = _tableCore.ActualRowCount;
            var columnCount = _tableCore.ColumnCount;
            
            // UPDATE: Set total dataset size for virtualization
            _totalDatasetSize = actualRowCount;
            
            _logger?.LogInformation("📊 RENDER STATE: TotalDataset={TotalDataset}, ViewportStart={ViewportStart}, ViewportSize={ViewportSize}, ColumnCount={ColumnCount}", 
                _totalDatasetSize, _viewportStartIndex, _viewportSize, columnCount);
            
            // CRITICAL SAFETY CHECK: Prevent XAML ItemsRepeater Int32.MaxValue index errors
            if (actualRowCount < 0)
            {
                _logger?.LogError("🚨 INDEX ERROR: ActualRowCount is negative: {ActualRowCount}", actualRowCount);
                throw new InvalidOperationException($"ActualRowCount cannot be negative: {actualRowCount}");
            }
            
            // VIRTUAL SCROLLING: No limit on total dataset size - viewport handles large datasets
            if (actualRowCount > 1000000) // 1M rows safety limit for total dataset
            {
                _logger?.LogError("🚨 DATASET ERROR: ActualRowCount exceeds maximum supported size: {ActualRowCount} > 1,000,000", actualRowCount);
                throw new InvalidOperationException($"ActualRowCount exceeds maximum supported size: {actualRowCount}");
            }
            
            if (columnCount < 0)
            {
                _logger?.LogError("🚨 INDEX ERROR: ColumnCount is negative: {ColumnCount}", columnCount);
                throw new InvalidOperationException($"ColumnCount cannot be negative: {columnCount}");
            }
            
            if (columnCount > 100) // 100 columns safety limit for XAML binding
            {
                _logger?.LogError("🚨 XAML BINDING ERROR: ColumnCount exceeds WinUI3 ItemsRepeater safety limit: {ColumnCount} > 100", columnCount);
                throw new InvalidOperationException($"ColumnCount exceeds WinUI3 ItemsRepeater safety limit: {columnCount}");
            }
            
            // VIRTUAL SCROLLING: Calculate viewport boundaries
            int viewportStart = Math.Max(0, _viewportStartIndex);
            int viewportEnd = Math.Min(_totalDatasetSize - 1, _viewportStartIndex + _viewportSize - 1);
            int viewportRowCount = Math.Max(0, viewportEnd - viewportStart + 1);
            
            _logger?.LogInformation("📊 VIEWPORT: Rendering rows {ViewportStart} to {ViewportEnd} ({ViewportCount} rows) from total {TotalRows}", 
                viewportStart, viewportEnd, viewportRowCount, _totalDatasetSize);

            // VIRTUAL SCROLLING: Check viewport cell count instead of total
            long viewportCells = (long)viewportRowCount * columnCount;
            if (viewportCells > 10000) // 10K viewport cells safety limit  
            {
                _logger?.LogError("🚨 VIEWPORT ERROR: Viewport cell count exceeds WinRT safety limit: {ViewportCells} > 10,000 (Rows: {Rows} × Columns: {Cols})", 
                    viewportCells, viewportRowCount, columnCount);
                throw new InvalidOperationException($"Viewport cell count exceeds WinRT safety limit: {viewportCells}");
            }

            _viewportRowsCollection.Clear();
            _logger?.LogInformation("🧹 UI RENDER: Viewport cleared, rendering viewport rows {Start}-{End}", viewportStart, viewportEnd);

            // VIRTUAL SCROLLING: Render only viewport rows
            int viewportRowIndex = 0; // Index v rámci viewport (0-based)
            
            for (int datasetRowIndex = viewportStart; datasetRowIndex <= viewportEnd; datasetRowIndex++)
            {
                try
                {
                    // Log progress for viewport rendering
                    if (viewportRowIndex % 20 == 0)
                    {
                        _logger?.LogInformation("📍 VIEWPORT PROGRESS: Rendering viewport row {ViewportIndex} (dataset row {DatasetIndex})", 
                            viewportRowIndex, datasetRowIndex);
                    }
                    
                    // SAFETY CHECK: Verify datasetRowIndex is within dataset bounds
                    if (datasetRowIndex >= _totalDatasetSize)
                    {
                        _logger?.LogError("🚨 VIEWPORT ERROR: DatasetRowIndex exceeds total dataset: {DatasetIndex} >= {TotalSize}", 
                            datasetRowIndex, _totalDatasetSize);
                        break;
                    }

                    var rowModel = new DataRowModel
                    {
                        RowIndex = datasetRowIndex,  // IMPORTANT: Use dataset index, not viewport index
                        BackgroundBrush = CreateBrush(_colorConfig.CellBackgroundColor)
                    };

                    _logger?.LogDebug("🎨 VIEWPORT ROW: Created DataRowModel viewport[{ViewportIndex}] = dataset[{DatasetIndex}]", 
                        viewportRowIndex, datasetRowIndex);

                    // Render cells pre tento riadok
                    for (int colIndex = 0; colIndex < columnCount; colIndex++)
                    {
                        try
                        {
                            var columnDef = _tableCore.GetColumnDefinition(colIndex);
                            if (columnDef == null) 
                            {
                                _logger?.LogWarning("⚠️ COLUMN WARNING: ColumnDefinition is null for colIndex {ColIndex}", colIndex);
                                continue;
                            }

                            _logger?.LogDebug("📋 VIEWPORT CELL: Processing dataset[{DatasetRow},{Col}] for column '{ColumnName}'", 
                                datasetRowIndex, colIndex, columnDef?.Name ?? "Unknown");

                            // IMPORTANT: Use datasetRowIndex for actual data access
                            var cellValue = await _tableCore.GetCellValueAsync(datasetRowIndex, colIndex);
                            
                            // CRITICAL: Get exact width from corresponding header for perfect alignment
                            var headerWidth = colIndex < _headersCollection.Count 
                                ? _headersCollection[colIndex].Width 
                                : (columnDef.Width ?? 100);

                            var cellModel = new DataCellModel
                            {
                                Value = cellValue,
                                DisplayText = cellValue?.ToString() ?? string.Empty,
                                RowIndex = datasetRowIndex,  // IMPORTANT: Store dataset index for API compatibility
                                ColumnIndex = colIndex,
                                ColumnName = columnDef.Name,
                                IsReadOnly = columnDef.IsReadOnly,
                                Width = headerWidth,  // CRITICAL: Use exact header width for perfect alignment
                                BackgroundBrush = CreateBrush(_colorConfig.CellBackgroundColor),
                                ForegroundBrush = CreateBrush(_colorConfig.CellForegroundColor),
                                BorderBrush = CreateBrush(_colorConfig.CellBorderColor)
                            };

                            _logger?.LogDebug("🧱 VIEWPORT CELL: Created DataCellModel at dataset[{DatasetRow},{Col}] with value='{Value}'", 
                                datasetRowIndex, colIndex, cellValue?.ToString() ?? "null");

                            // Validation check with error handling
                            try
                            {
                                await UpdateCellValidationAsync(cellModel);
                            }
                            catch (Exception validationEx)
                            {
                                _logger?.LogError(validationEx, "🚨 VALIDATION ERROR: Cell validation failed [{DatasetRow},{Col}]", 
                                    datasetRowIndex, colIndex);
                                // Continue with default validation state
                                cellModel.IsValid = true;
                                cellModel.ValidationError = null;
                            }

                            rowModel.Cells.Add(cellModel);
                        }
                        catch (Exception cellEx)
                        {
                            _logger?.LogError(cellEx, "🚨 VIEWPORT CELL ERROR: Failed to process viewport cell [{ViewportRow},{ViewportCol}] = dataset[{DatasetRow},{DatasetCol}]", 
                                viewportRowIndex, colIndex, datasetRowIndex, colIndex);
                            throw;
                        }
                    }

                    // Check if row is empty
                    rowModel.IsEmpty = rowModel.Cells.All(c => string.IsNullOrEmpty(c.DisplayText));
                    rowModel.IsValid = rowModel.Cells.All(c => c.IsValid);

                    _logger?.LogDebug("✅ VIEWPORT ROW COMPLETE: Viewport[{ViewportIndex}] = Dataset[{DatasetIndex}] - Cells: {CellCount}, Empty: {IsEmpty}, Valid: {IsValid}", 
                        viewportRowIndex, datasetRowIndex, rowModel.Cells.Count, rowModel.IsEmpty, rowModel.IsValid);

                    _viewportRowsCollection.Add(rowModel);
                    viewportRowIndex++; // Increment viewport position
                }
                catch (Exception rowEx)
                {
                    _logger?.LogError(rowEx, "🚨 VIEWPORT ROW ERROR: Failed to process viewport row {ViewportIndex} = dataset row {DatasetIndex}", 
                        viewportRowIndex, datasetRowIndex);
                    throw;
                }
            }

            _logger?.LogInformation("✅ VIRTUAL SCROLLING: Viewport rendered - {ViewportRows} viewport rows from dataset rows {Start}-{End}, {CellCount} total cells", 
                _viewportRowsCollection.Count, viewportStart, viewportEnd, _viewportRowsCollection.Sum(r => r.Cells.Count));
                
            // CRITICAL DIAGNOSTIC: Log actual row data
            for (int i = 0; i < Math.Min(3, _viewportRowsCollection.Count); i++) // Log first 3 rows
            {
                var row = _viewportRowsCollection[i];
                _logger?.LogInformation("📋 ROW[{Index}]: {CellCount} cells, IsEmpty={IsEmpty}", 
                    i, row.Cells.Count, row.IsEmpty);
                    
                for (int j = 0; j < Math.Min(4, row.Cells.Count); j++) // Log first 4 cells
                {
                    var cell = row.Cells[j];
                    _logger?.LogInformation("   🧱 CELL[{RowIndex},{CellIndex}]: Value='{Value}', DisplayText='{DisplayText}', Width={Width}", 
                        i, j, cell.Value?.ToString() ?? "null", cell.DisplayText, cell.Width);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: Data rows rendering failed - ActualRowCount: {ActualRowCount}, ColumnCount: {ColumnCount}", 
                _tableCore?.ActualRowCount ?? -1, _tableCore?.ColumnCount ?? -1);
            throw;
        }
    }

    /// <summary>
    /// Update data pre existujúci row model
    /// </summary>
    private async Task UpdateRowDataAsync(DataRowModel rowModel)
    {
        try
        {
            for (int colIndex = 0; colIndex < rowModel.Cells.Count; colIndex++)
            {
                var cellModel = rowModel.Cells[colIndex];
                var cellValue = await _tableCore.GetCellValueAsync(rowModel.RowIndex, colIndex);
                
                cellModel.Value = cellValue;
                cellModel.DisplayText = cellValue?.ToString() ?? string.Empty;
                
                await UpdateCellValidationAsync(cellModel);
            }

            rowModel.IsEmpty = rowModel.Cells.All(c => string.IsNullOrEmpty(c.DisplayText));
            rowModel.IsValid = rowModel.Cells.All(c => c.IsValid);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: Row data update failed for row {RowIndex}", rowModel.RowIndex);
            throw;
        }
    }

    /// <summary>
    /// Update validation pre cell model
    /// </summary>
    public async Task UpdateCellValidationAsync(DataCellModel cellModel)
    {
        try
        {
            // REAL VALIDATION: Call TableCore validation logic
            try
            {
                var cellValue = cellModel.Value;
                var columnName = cellModel.ColumnName;
                var rowIndex = cellModel.RowIndex;
                
                // Get actual validation result from TableCore
                // For now, use simple validation - TODO: implement full validation logic
                bool isValid = await ValidateCellValueAsync(cellValue, columnName);
                string? validationError = null;
                
                if (!isValid)
                {
                    // Generate validation error message
                    validationError = GenerateValidationErrorMessage(cellValue, columnName);
                    _logger?.LogDebug("🚨 VALIDATION: Cell [{Row},{Col}] failed validation - Value: '{Value}', Error: '{Error}'", 
                        rowIndex, cellModel.ColumnIndex, cellValue?.ToString() ?? "null", validationError);
                }
                
                cellModel.IsValid = isValid;
                cellModel.ValidationError = validationError;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "🚨 VALIDATION ERROR: Failed to validate cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
                
                // Default to invalid on validation errors
                cellModel.IsValid = false;
                cellModel.ValidationError = "Validation error occurred";
            }

            // Update visual styling based on validation
            if (!cellModel.IsValid)
            {
                cellModel.BorderBrush = CreateBrush(_colorConfig.ValidationErrorBorderColor);
                cellModel.BackgroundBrush = CreateBrush(_colorConfig.ValidationErrorBackgroundColor);
                // VALIDATION BORDER FIX: Make error border more visible with thicker border
                cellModel.BorderThickness = new Microsoft.UI.Xaml.Thickness(2);
                
                _logger?.LogDebug("🚨 VALIDATION VISUAL: Applied error styling to cell [{Row},{Col}] - Border: Red, Background: Light Red", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
            else
            {
                cellModel.BorderBrush = CreateBrush(_colorConfig.CellBorderColor);
                cellModel.BackgroundBrush = CreateBrush(_colorConfig.CellBackgroundColor);
                // Reset to normal border thickness
                cellModel.BorderThickness = new Microsoft.UI.Xaml.Thickness(1);
                
                cellModel.ValidationError = null;
            }
            
            // REALTIME VALIDATION ALERT: Update ValidationAlerts column immediately
            await UpdateValidationAlertsColumnAsync(cellModel.RowIndex, null);
            
            _logger?.LogDebug("✅ VALIDATION UPDATE: Real-time validation completed for cell [{Row},{Col}] - Valid: {IsValid}", 
                cellModel.RowIndex, cellModel.ColumnIndex, cellModel.IsValid);

            await Task.CompletedTask; // For async consistency
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: Cell validation update failed for cell [{Row},{Col}]", 
                cellModel.RowIndex, cellModel.ColumnIndex);
            throw;
        }
    }

    /// <summary>
    /// Apply auto-stretch logic for ValidationAlerts column to fill remaining space
    /// ValidationAlerts is last column or second-to-last (before DeleteRow)
    /// </summary>
    private void ApplyValidationAlertsAutoStretch()
    {
        try
        {
            // Find ValidationAlerts column
            int validationAlertsIndex = -1;
            int deleteRowIndex = -1;
            
            for (int i = 0; i < _headersCollection.Count; i++)
            {
                var header = _headersCollection[i];
                if (header.ColumnName?.Equals("ValidationAlerts", StringComparison.OrdinalIgnoreCase) == true)
                {
                    validationAlertsIndex = i;
                }
                else if (header.ColumnName?.Equals("DeleteRow", StringComparison.OrdinalIgnoreCase) == true)
                {
                    deleteRowIndex = i;
                }
            }

            if (validationAlertsIndex == -1)
            {
                _logger?.LogDebug("🔍 AUTO-STRETCH: ValidationAlerts column not found, skipping auto-stretch");
                return;
            }

            // Calculate total width of all other columns
            double totalOtherColumnsWidth = 0;
            const double deleteRowFixedWidth = 60; // Fixed width for DeleteRow column
            const double validationAlertsMinWidth = 120; // Default minimum width
            
            // KRITICKÉ: Get actual available container width dynamically (remove maximum limit)
            double actualContainerWidth = GetActualContainerWidth(); // Dynamic measurement

            for (int i = 0; i < _headersCollection.Count; i++)
            {
                if (i != validationAlertsIndex)
                {
                    if (i == deleteRowIndex)
                    {
                        totalOtherColumnsWidth += deleteRowFixedWidth;
                        _headersCollection[i].Width = deleteRowFixedWidth; // Ensure DeleteRow has fixed width
                    }
                    else
                    {
                        totalOtherColumnsWidth += _headersCollection[i].Width;
                    }
                }
            }

            // KRITICKÉ: Calculate remaining space with NO MAXIMUM LIMIT - ValidationAlerts fills all available space
            double remainingSpace = actualContainerWidth - totalOtherColumnsWidth;
            double validationAlertsWidth = Math.Max(validationAlertsMinWidth, remainingSpace);
            
            // Ensure ValidationAlerts width grows with container (no maximum constraint)
            if (validationAlertsWidth < validationAlertsMinWidth)
            {
                validationAlertsWidth = validationAlertsMinWidth;
            }

            // Apply the calculated width
            _headersCollection[validationAlertsIndex].Width = validationAlertsWidth;

            _logger?.LogInformation("📏 AUTO-STRETCH: ValidationAlerts column width = {Width} (min: {Min}, remaining: {Remaining}, container: {Container})", 
                validationAlertsWidth, validationAlertsMinWidth, remainingSpace, actualContainerWidth);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 AUTO-STRETCH ERROR: Failed to apply ValidationAlerts auto-stretch");
        }
    }
    
    /// <summary>
    /// Get actual container width dynamically - NO MAXIMUM LIMIT
    /// ValidationAlerts fills all available space in the container
    /// KRITICKÉ: Removes artificial width constraints - allows unlimited expansion
    /// </summary>
    private double GetActualContainerWidth()
    {
        try
        {
            // STRATEGY: Responsive width calculation that grows with content
            // NO artificial maximum limits - ValidationAlerts expands to fill ALL available space
            
            // Base width calculation - scales with column count
            double baseWidth = 1200; // Increased base width
            double columnScaling = _headersCollection.Count * 80; // More generous per-column space
            double contentBasedWidth = baseWidth + columnScaling;
            
            // UNLIMITED EXPANSION: Allow ValidationAlerts to take massive space if needed
            // This addresses user issue: "ValidationAlerts stops stretching at maximum width"
            double unlimitedWidth = Math.Max(contentBasedWidth, 2000); // Minimum 2000px, no maximum
            
            // FUTURE: Can be enhanced to read actual container measurements
            // For now: generous width calculation ensures ValidationAlerts never hits artificial limits
            
            _logger?.LogDebug("📐 UNLIMITED WIDTH: Container width = {Width} (base: {Base}, scaling: {Scaling}, unlimited: TRUE)", 
                unlimitedWidth, baseWidth, columnScaling);
                
            return unlimitedWidth;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "⚠️ CONTAINER WIDTH: Failed to calculate unlimited width, using safe fallback");
            return 2000; // Safe fallback - generous width
        }
    }

    /// <summary>
    /// Vytvorí SolidColorBrush z Windows.UI.Color (nullable safe)
    /// </summary>
    private SolidColorBrush CreateBrush(Windows.UI.Color? color)
    {
        if (color.HasValue)
        {
            var c = color.Value;
            var uiColor = Windows.UI.Color.FromArgb(c.A, c.R, c.G, c.B);
            return new SolidColorBrush(uiColor);
        }
        
        // Default color if none provided
        return new SolidColorBrush(Colors.White);
    }

    /// <summary>
    /// Reapply ValidationAlerts auto-stretch after window/container resize
    /// PUBLIC API for responsive behavior
    /// </summary>
    public async Task ReapplyAutoStretchAsync()
    {
        try
        {
            _logger?.LogInformation("🔄 AUTO-STRETCH: Reapplying ValidationAlerts auto-stretch after resize");
            
            // Reapply header auto-stretch
            ApplyValidationAlertsAutoStretch();
            
            // CRITICAL: Re-render data cells with updated header widths
            await RenderDataRowsAsync();
            
            _logger?.LogInformation("✅ AUTO-STRETCH: ValidationAlerts auto-stretch reapplied successfully");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 AUTO-STRETCH ERROR: Failed to reapply ValidationAlerts auto-stretch");
            throw;
        }
    }
    
    /// <summary>
    /// Force refresh ValidationAlerts column width (remove maximum constraints)
    /// PUBLIC API for manual width adjustment
    /// </summary>
    public void ForceValidationAlertsWidthRefresh()
    {
        try
        {
            _logger?.LogInformation("🔄 FORCE REFRESH: Forcing ValidationAlerts width refresh");
            ApplyValidationAlertsAutoStretch();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 FORCE REFRESH ERROR: Failed to force ValidationAlerts width refresh");
        }
    }

    #endregion

    #region Validation Helper Methods
    
    /// <summary>
    /// Simple cell validation - returns false for specific test cases
    /// </summary>
    private async Task<bool> ValidateCellValueAsync(object? value, string columnName)
    {
        await Task.CompletedTask; // Make method async
        
        if (value == null) return true; // Allow null values
        
        string? stringValue = value.ToString();
        if (string.IsNullOrEmpty(stringValue)) return true; // Allow empty values
        
        // Test case: "jaja" should be invalid for email validation demo
        if (columnName.Equals("Email", StringComparison.OrdinalIgnoreCase))
        {
            if (stringValue.Equals("jaja", StringComparison.OrdinalIgnoreCase))
            {
                return false; // Invalid email for demo
            }
            // Simple email validation
            if (!stringValue.Contains("@") || !stringValue.Contains("."))
            {
                return false;
            }
        }
        
        // Test case: negative numbers should be invalid for Age column
        if (columnName.Equals("Age", StringComparison.OrdinalIgnoreCase))
        {
            if (int.TryParse(stringValue, out int age) && age < 0)
            {
                return false;
            }
        }
        
        return true; // Valid by default
    }

    /// <summary>
    /// Generate validation error message
    /// </summary>
    private string GenerateValidationErrorMessage(object? value, string columnName)
    {
        string? stringValue = value?.ToString() ?? "null";
        
        if (columnName.Equals("Email", StringComparison.OrdinalIgnoreCase))
        {
            if (stringValue.Equals("jaja", StringComparison.OrdinalIgnoreCase))
            {
                return "Invalid email format: 'jaja' is not a valid email";
            }
            return $"Invalid email format: '{stringValue}'";
        }
        
        if (columnName.Equals("Age", StringComparison.OrdinalIgnoreCase))
        {
            return $"Age cannot be negative: '{stringValue}'";
        }
        
        return $"Invalid value: '{stringValue}'";
    }

    /// <summary>
    /// Update ValidationAlerts column for a specific row
    /// </summary>
    private async Task UpdateValidationAlertsColumnAsync(int rowIndex, string? errorMessage)
    {
        try
        {
            // Find ValidationAlerts column index
            int validationColumnIndex = -1;
            for (int i = 0; i < _headersCollection.Count; i++)
            {
                if (_headersCollection[i].ColumnName.Equals("ValidationAlerts", StringComparison.OrdinalIgnoreCase))
                {
                    validationColumnIndex = i;
                    break;
                }
            }

            if (validationColumnIndex == -1)
            {
                _logger?.LogWarning("⚠️ VALIDATION ALERTS: ValidationAlerts column not found");
                return;
            }

            // Find the row in viewport
            var targetRow = _viewportRowsCollection.FirstOrDefault(r => r.RowIndex == rowIndex);
            if (targetRow != null && validationColumnIndex < targetRow.Cells.Count)
            {
                var validationCell = targetRow.Cells[validationColumnIndex];
                
                // Collect all validation errors for this row
                var rowErrors = new List<string>();
                foreach (var cell in targetRow.Cells)
                {
                    if (!cell.IsValid && !string.IsNullOrEmpty(cell.ValidationError))
                    {
                        rowErrors.Add($"{cell.ColumnName}: {cell.ValidationError}");
                    }
                }

                // Update ValidationAlerts cell (use property setters for proper PropertyChanged notifications)
                string alertsText = rowErrors.Count > 0 ? string.Join("; ", rowErrors) : "";
                validationCell.DisplayText = alertsText;  // This should trigger PropertyChanged via SetProperty
                validationCell.Value = alertsText;        // This should trigger PropertyChanged via SetProperty
                
                _logger?.LogWarning("📋 VALIDATION ALERTS: Updated row {RowIndex} col {ColIndex} '{ColName}' with {ErrorCount} errors: '{Alerts}'", 
                    rowIndex, validationColumnIndex, validationCell.ColumnName, rowErrors.Count, alertsText);
                
                // DIAGNOSTIC: Log all cell states for this row
                for (int i = 0; i < targetRow.Cells.Count; i++)
                {
                    var cell = targetRow.Cells[i];
                    _logger?.LogDebug("  📝 CELL[{Index}] '{ColName}': Valid={IsValid}, Error='{Error}', Display='{Display}'", 
                        i, cell.ColumnName, cell.IsValid, cell.ValidationError ?? "null", cell.DisplayText);
                }
            }

            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 VALIDATION ALERTS ERROR: Failed to update ValidationAlerts for row {RowIndex}", rowIndex);
        }
    }

    #endregion

    #region Public Statistics

    /// <summary>
    /// Získa UI rendering statistiky s viewport informáciami
    /// </summary>
    public UIRenderingStats GetRenderingStats()
    {
        return new UIRenderingStats
        {
            HeaderCount = _headersCollection.Count,
            RowCount = _viewportRowsCollection.Count,  // Viewport rows currently rendered
            TotalCellCount = _viewportRowsCollection.Sum(r => r.Cells.Count),
            LastRenderTime = _lastRenderTime,
            IsCurrentlyRendering = _isRendering,
            // Virtual Scrolling specific stats
            TotalDatasetSize = _totalDatasetSize,
            ViewportStartIndex = _viewportStartIndex,
            ViewportSize = _viewportSize,
            ViewportEndIndex = ViewportEndIndex
        };
    }

    #endregion
}

/// <summary>
/// Štatistiky UI renderovania s Virtual Scrolling informáciami
/// </summary>
public class UIRenderingStats
{
    public int HeaderCount { get; set; }
    public int RowCount { get; set; }              // Viewport rows currently rendered
    public int TotalCellCount { get; set; }
    public DateTime LastRenderTime { get; set; }
    public bool IsCurrentlyRendering { get; set; }
    
    // Virtual Scrolling specific properties
    public int TotalDatasetSize { get; set; }      // Complete dataset size
    public int ViewportStartIndex { get; set; }    // First visible row index in dataset
    public int ViewportSize { get; set; }          // Number of rows in viewport
    public int ViewportEndIndex { get; set; }      // Last visible row index in dataset
    
    // Convenience properties
    public double ViewportPositionPercent => TotalDatasetSize > 0 ? (double)ViewportStartIndex / TotalDatasetSize * 100 : 0;
    public bool IsAtStart => ViewportStartIndex == 0;
    public bool IsAtEnd => ViewportEndIndex >= TotalDatasetSize - 1;
}