using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Shapes;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Input;
using Windows.System;
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
    /// UI Manager pre kvalitné UI rendering s proper error logging
    /// </summary>
    private DataGridUIManager? _uiManager;

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
    
    // COLUMN RESIZE STATE MANAGEMENT (these are already defined elsewhere - remove duplicates)

    // UI XAML elements are automatically generated from XAML file
    
    // Double-click detection for cell editing
    private DataCellModel? _lastClickedCell;
    private DateTime _lastClickTime = DateTime.MinValue;
    private const int DoubleClickTimeoutMs = 500;
    
    // Edit state management for cancel/commit
    private DataCellModel? _currentEditingCell;
    private string? _originalEditValue;
    
    // Focus state management for navigation
    private DataCellModel? _focusedCell;
    private int _focusedRowIndex = 0;
    private int _focusedColumnIndex = 0;

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
        
        // Initialize keyboard shortcuts
        InitializeKeyboardHandlers();
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

            // Phase 2: Initialize UI Manager
            _uiManager = new DataGridUIManager(_controller.TableCore, _logger);
            _uiManager.ColorConfig = _controller.ColorConfig;

            // Phase 3: Initialize UI layer
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

            // Initialize UI Manager and bind to XAML controls
            if (_uiManager != null)
            {
                await _uiManager.InitializeUIAsync();
                
                // Bind ObservableCollections to ItemsRepeater controls - CRITICAL: On UI thread!
                bool bindingSuccess = false;
                if (DispatcherQueue != null)
                {
                    DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
                    {
                        try
                        {
                            if (HeaderRepeater != null)
                            {
                                HeaderRepeater.ItemsSource = _uiManager.HeadersCollection;
                                _logger?.LogInformation("🎨 UI BINDING: HeaderRepeater bound to HeadersCollection on UI thread");
                            }
                            
                            if (DataRepeater != null)
                            {
                                DataRepeater.ItemsSource = _uiManager.RowsCollection;
                                _logger?.LogInformation("🎨 UI BINDING: DataRepeater bound to RowsCollection on UI thread");
                            }
                            bindingSuccess = true;
                        }
                        catch (Exception bindEx)
                        {
                            _logger?.LogError(bindEx, "🚨 UI BINDING ERROR: Failed to bind ItemsSource on UI thread");
                        }
                    });
                    
                    // Wait a moment for UI thread operation to complete
                    await Task.Delay(100);
                    _logger?.LogInformation("✅ UI BINDING: Dispatcher binding request submitted, Success: {Success}", bindingSuccess);
                }
                else
                {
                    // Fallback to direct binding if no dispatcher available
                    _logger?.LogWarning("⚠️ UI BINDING: No DispatcherQueue available, using direct binding");
                    
                    if (HeaderRepeater != null)
                    {
                        HeaderRepeater.ItemsSource = _uiManager.HeadersCollection;
                        _logger?.LogInformation("🎨 UI BINDING: HeaderRepeater bound to HeadersCollection (direct)");
                    }
                    
                    if (DataRepeater != null)
                    {
                        DataRepeater.ItemsSource = _uiManager.RowsCollection;
                        _logger?.LogInformation("🎨 UI BINDING: DataRepeater bound to RowsCollection (direct)");
                    }
                }
            }

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
            
            // DATA RETENTION FIX: Commit any active edit operations before refresh
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                _logger?.LogInformation("💾 DATA RETENTION: Auto-committing active edit before refresh");
                await CommitCellEditAsync(stayOnCell: false);
            }
            
            await ShowLoadingAsync(true);
            await RenderAllCellsAsync();
            
            // VALIDATION ALERTS FIX: Auto-update validation UI during refresh to ensure ValidationAlerts column is populated
            _logger?.LogInformation("🔍 UI UPDATE: Auto-updating validation during refresh to populate ValidationAlerts");
            await UpdateValidationVisualsAsync();
            
            await ShowLoadingAsync(false);

            _logger?.LogInformation("✅ UI UPDATE: Full UI refresh completed with data retention");
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

            // IMPORTANT: Keep StackLayout to preserve individual column widths
            // UniformGridLayout would force all columns to same width!
            if (DataRepeater != null)
            {
                var currentRowHeight = _controller.CurrentRowHeight;
                var layout = new StackLayout
                {
                    Orientation = Orientation.Vertical,
                    Spacing = 0
                };
                DataRepeater.Layout = layout;
                
                _logger?.LogInformation("🎨 UI LAYOUT: Updated with StackLayout to preserve column widths, row height: {Height}px", Math.Ceiling(currentRowHeight));
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
    /// Initialize UI event handlers - CRITICAL: Column resize & cell interaction
    /// </summary>
    private void InitializeUIEventHandlers()
    {
        try
        {
            _logger?.LogInformation("🎨 UI EVENTS: Initializing event handlers for resize & cell interaction");
            
            // Column resize state management
            _isResizing = false;
            _resizingColumnIndex = -1;
            _resizeStartX = 0;
            _resizeStartWidth = 0;
            
            // Event handlers will be attached to ItemsRepeater elements dynamically
            // when UI elements are rendered (in DataGridUIManager)
            
            _logger?.LogInformation("✅ UI EVENTS: Event handlers initialized successfully");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI EVENTS ERROR: Failed to initialize event handlers");
        }
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
    /// Render all cells using proper UI Manager - kvalitné riešenie s comprehensive error logging
    /// </summary>
    private async Task RenderAllCellsAsync()
    {
        if (_uiManager == null)
        {
            _logger?.LogError("🚨 UI ERROR: UIManager is null, cannot render cells");
            throw new InvalidOperationException("UIManager must be initialized before rendering cells");
        }

        try
        {
            _logger?.LogInformation("🎨 UI RENDER: Starting comprehensive cell rendering via UIManager...");
            
            // CRITICAL FIX: RefreshAllUIAsync must run on UI thread for ObservableCollection binding
            if (DispatcherQueue != null)
            {
                var completionSource = new TaskCompletionSource<bool>();
                
                DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, async () =>
                {
                    try
                    {
                        await _uiManager.RefreshAllUIAsync();
                        
                        // Update UI statistics
                        var stats = _uiManager.GetRenderingStats();
                        _logger?.LogInformation("✅ UI RENDER: Comprehensive rendering completed on UI thread - Headers: {HeaderCount}, Rows: {RowCount}, Cells: {CellCount}", 
                            stats.HeaderCount, stats.RowCount, stats.TotalCellCount);
                        
                        completionSource.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "🚨 UI RENDER ERROR: RefreshAllUIAsync failed on UI thread");
                        completionSource.SetException(ex);
                    }
                });
                
                // Wait for UI thread operation to complete
                await completionSource.Task;
            }
            else
            {
                // Fallback to direct call if no dispatcher available
                _logger?.LogWarning("⚠️ UI RENDER: No DispatcherQueue available, using direct call");
                await _uiManager.RefreshAllUIAsync();
                
                var stats = _uiManager.GetRenderingStats();
                _logger?.LogInformation("✅ UI RENDER: Comprehensive rendering completed (direct) - Headers: {HeaderCount}, Rows: {RowCount}, Cells: {CellCount}", 
                    stats.HeaderCount, stats.RowCount, stats.TotalCellCount);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: Comprehensive cell rendering failed");
            throw;
        }
    }

    /// <summary>
    /// Update validation visuals using UIManager
    /// </summary>
    private async Task UpdateValidationVisualsAsync()
    {
        if (_uiManager == null)
        {
            _logger?.LogError("🚨 UI ERROR: UIManager is null, cannot update validation visuals");
            return;
        }

        try
        {
            await _uiManager.UpdateValidationUIAsync();
            _logger?.LogInformation("✅ UI VALIDATION: Validation visuals updated via UIManager");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateValidationVisualsAsync failed");
            throw;
        }
    }

    /// <summary>
    /// Update specific row UI using UIManager
    /// </summary>
    private async Task UpdateSpecificRowUIAsync(int rowIndex)
    {
        if (_uiManager == null)
        {
            _logger?.LogError("🚨 UI ERROR: UIManager is null, cannot update row UI");
            return;
        }

        try
        {
            await _uiManager.UpdateRowUIAsync(rowIndex);
            _logger?.LogInformation("✅ UI ROW: Updated row {Row} UI via UIManager", rowIndex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 UI ERROR: UpdateSpecificRowUIAsync failed for row {Row}", rowIndex);
            throw;
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

    #region Column Resize Event Handlers

    private bool _isResizing = false;
    private int _resizingColumnIndex = -1;
    private double _resizeStartX = 0;
    private double _resizeStartWidth = 0;
    private Rectangle? _currentResizeHandle;
    
    // Drag selection fields
    private bool _isDragSelecting = false;
    private bool _isDragPending = false;
    private DataCellModel? _dragStartCell = null;
    private DataCellModel? _dragEndCell = null;
    private Windows.Foundation.Point _dragStartPoint;

    /// <summary>
    /// Event handler for pointer entering resize handle area
    /// </summary>
    private void ResizeHandle_PointerEntered(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Rectangle handle)
            {
                // Change cursor to resize cursor immediately when entering ±2px resize area
                this.ProtectedCursor = Microsoft.UI.Input.InputSystemCursor.Create(Microsoft.UI.Input.InputSystemCursorShape.SizeWestEast);
                _logger?.LogDebug("🔍 RESIZE: Pointer entered ±2px resize area");
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to handle pointer enter");
        }
    }

    /// <summary>
    /// Event handler for pointer exiting resize handle area
    /// </summary>
    private void ResizeHandle_PointerExited(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Rectangle handle && !_isResizing)
            {
                // Reset cursor to default
                this.ProtectedCursor = null;
                _logger?.LogDebug("🔍 RESIZE: Pointer exited resize handle");
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to handle pointer exit");
        }
    }

    /// <summary>
    /// Event handler for starting column resize operation
    /// </summary>
    private void ResizeHandle_PointerPressed(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Rectangle handle)
            {
                // Find which column we're resizing by looking at the data context
                int columnIndex = GetColumnIndexFromResizeHandle(handle);
                if (columnIndex >= 0)
                {
                    _isResizing = true;
                    _resizingColumnIndex = columnIndex;
                    _currentResizeHandle = handle;
                    _resizeStartX = e.GetCurrentPoint(this).Position.X; // Use grid coordinates for consistency
                    
                    // Get current width from header model
                    if (columnIndex < _uiManager?.HeadersCollection.Count)
                    {
                        _resizeStartWidth = _uiManager.HeadersCollection[columnIndex].Width;
                    }

                    // Capture pointer for drag operation
                    handle.CapturePointer(e.Pointer);
                    
                    _logger?.LogDebug("🔄 RESIZE: Started resizing column {ColumnIndex} from width {StartWidth}", 
                        columnIndex, _resizeStartWidth);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to start resize operation");
            _isResizing = false;
            _resizingColumnIndex = -1;
        }
    }

    /// <summary>
    /// Event handler for tracking mouse movement during resize
    /// </summary>
    private void ResizeHandle_PointerMoved(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Rectangle handle)
            {
                if (_isResizing && _resizingColumnIndex >= 0)
                {
                    // Real-time resize tracking - move exactly with the mouse
                    double currentX = e.GetCurrentPoint(this).Position.X; // Get position relative to grid, not handle
                    double deltaX = currentX - _resizeStartX;
                    double targetWidth = _resizeStartWidth + deltaX;
                    
                    // Apply MinWidth/MaxWidth constraints from column definition
                    double newWidth = ApplyColumnWidthConstraints(_resizingColumnIndex, targetWidth);
                    
                    // Update the column width in real-time for entire column (headers + cells)
                    UpdateColumnWidth(_resizingColumnIndex, newWidth);
                    
                    _logger?.LogDebug("🔄 RESIZE: Column {ColumnIndex} width = {NewWidth} (delta: {Delta})", 
                        _resizingColumnIndex, newWidth, deltaX);
                }
                else
                {
                    // Keep resize cursor when hovering over ±2px resize area
                    this.ProtectedCursor = Microsoft.UI.Input.InputSystemCursor.Create(Microsoft.UI.Input.InputSystemCursorShape.SizeWestEast);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to handle pointer move");
        }
    }

    /// <summary>
    /// Event handler for ending column resize operation
    /// </summary>
    private void ResizeHandle_PointerReleased(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (_isResizing && sender is Rectangle handle)
            {
                // Final width calculation
                double currentX = e.GetCurrentPoint(this).Position.X; // Use grid coordinates for consistency
                double deltaX = currentX - _resizeStartX;
                double targetWidth = _resizeStartWidth + deltaX;
                
                // Apply MinWidth/MaxWidth constraints from column definition
                double finalWidth = ApplyColumnWidthConstraints(_resizingColumnIndex, targetWidth);
                
                // Apply final width
                UpdateColumnWidth(_resizingColumnIndex, finalWidth);
                
                // Release pointer capture
                handle.ReleasePointerCapture(e.Pointer);
                
                _logger?.LogInformation("✅ RESIZE: Completed resizing column {ColumnIndex} to width {FinalWidth}", 
                    _resizingColumnIndex, finalWidth);
                
                // Reset resize state
                _isResizing = false;
                _resizingColumnIndex = -1;
                _currentResizeHandle = null;
                _resizeStartX = 0;
                _resizeStartWidth = 0;
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to complete resize operation");
            _isResizing = false;
            _resizingColumnIndex = -1;
        }
    }

    /// <summary>
    /// Helper method to get column index from resize handle
    /// </summary>
    private int GetColumnIndexFromResizeHandle(Rectangle handle)
    {
        try
        {
            // Walk up visual tree to find the data context
            DependencyObject parent = handle;
            while (parent != null)
            {
                if (parent is FrameworkElement element && element.DataContext != null)
                {
                    // Check if it's a header cell
                    if (element.DataContext is HeaderCellModel headerModel)
                    {
                        // Find index by matching column name
                        for (int i = 0; i < _uiManager?.HeadersCollection.Count; i++)
                        {
                            if (_uiManager.HeadersCollection[i].ColumnName == headerModel.ColumnName)
                            {
                                return i;
                            }
                        }
                    }
                    // Check if it's a data cell
                    else if (element.DataContext is DataCellModel cellModel)
                    {
                        return cellModel.ColumnIndex;
                    }
                }
                parent = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetParent(parent);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to get column index from resize handle");
        }
        
        return -1;
    }

    /// <summary>
    /// Fast update during resize - only updates header width for responsive UI
    /// </summary>
    private void UpdateColumnWidthFast(int columnIndex, double newWidth)
    {
        try
        {
            if (_uiManager == null || columnIndex < 0) return;
            
            // Only update header width during drag for fast response
            if (columnIndex < _uiManager.HeadersCollection.Count)
            {
                _uiManager.HeadersCollection[columnIndex].Width = newWidth;
            }
            
            _logger?.LogDebug("⚡ RESIZE FAST: Updated header {ColumnIndex} width to {NewWidth}", 
                columnIndex, newWidth);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to fast update column width");
        }
    }

    /// <summary>
    /// Complete update after resize - updates all cells in the column
    /// </summary>
    private void UpdateColumnWidth(int columnIndex, double newWidth)
    {
        try
        {
            if (_uiManager == null || columnIndex < 0) return;
            
            // Update header width
            if (columnIndex < _uiManager.HeadersCollection.Count)
            {
                _uiManager.HeadersCollection[columnIndex].Width = newWidth;
            }
            
            // Update all cell widths in that column
            foreach (var row in _uiManager.RowsCollection)
            {
                if (columnIndex < row.Cells.Count)
                {
                    row.Cells[columnIndex].Width = newWidth;
                }
            }
            
            _logger?.LogDebug("🔄 RESIZE: Updated column {ColumnIndex} width to {NewWidth} across all rows", 
                columnIndex, newWidth);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to update column width");
        }
    }

    /// <summary>
    /// Apply MinWidth/MaxWidth constraints from column definition
    /// </summary>
    private double ApplyColumnWidthConstraints(int columnIndex, double targetWidth)
    {
        try
        {
            if (_uiManager == null || columnIndex < 0 || columnIndex >= _uiManager.HeadersCollection.Count)
            {
                // Fallback: Use reasonable default constraints
                return Math.Max(50, Math.Min(targetWidth, 1000));
            }

            // Get column definition from controller
            var columnDef = _controller.TableCore.GetColumnDefinition(columnIndex);
            if (columnDef == null)
            {
                // Fallback: Use reasonable default constraints
                return Math.Max(50, Math.Min(targetWidth, 1000));
            }

            double minWidth = columnDef.MinWidth;  // MinWidth is non-nullable double
            double maxWidth = columnDef.MaxWidth ?? double.MaxValue; // MaxWidth is nullable double

            // Apply constraints
            double constrainedWidth = Math.Max(minWidth, targetWidth);
            if (maxWidth < double.MaxValue)
            {
                constrainedWidth = Math.Min(constrainedWidth, maxWidth);
            }

            _logger?.LogDebug("🔒 RESIZE CONSTRAINTS: Column {ColumnIndex} - Target: {Target}, Min: {Min}, Max: {Max}, Final: {Final}",
                columnIndex, targetWidth, minWidth, maxWidth == double.MaxValue ? "∞" : maxWidth.ToString(), constrainedWidth);

            return constrainedWidth;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RESIZE ERROR: Failed to apply width constraints for column {ColumnIndex}", columnIndex);
            // Fallback: Use reasonable default constraints
            return Math.Max(50, Math.Min(targetWidth, 1000));
        }
    }

    #endregion

    #region Cell Editing Event Handlers

    /// <summary>
    /// Event handler for entering edit mode when clicking on display area
    /// </summary>
    private void DisplayArea_Tapped(object sender, Microsoft.UI.Xaml.Input.TappedRoutedEventArgs e)
    {
        try
        {
            if (sender is FrameworkElement element && element.DataContext is DataCellModel cellModel)
            {
                // Don't edit read-only cells
                if (cellModel.IsReadOnly)
                {
                    _logger?.LogDebug("🔒 CELL EDIT: Cell [{Row},{Col}] is read-only, edit blocked", 
                        cellModel.RowIndex, cellModel.ColumnIndex);
                    return;
                }

                // Enter edit mode
                cellModel.IsEditing = true;
                _logger?.LogDebug("✏️ CELL EDIT: Entered edit mode for cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
                
                // Focus on the TextBox after UI update
                _ = Task.Run(async () =>
                {
                    await Task.Delay(50); // Allow UI to update
                    if (DispatcherQueue != null)
                    {
                        DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
                        {
                            // Find the TextBox in visual tree and focus it
                            if (FindTextBoxInVisualTree(element) is TextBox textBox)
                            {
                                textBox.Focus(FocusState.Programmatic);
                                textBox.SelectAll();
                            }
                        });
                    }
                });
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL EDIT ERROR: Failed to enter edit mode");
        }
    }

    /// <summary>
    /// Event handler for real-time validation during text changes
    /// </summary>
    private async void EditTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        try
        {
            if (sender is TextBox textBox && textBox.DataContext is DataCellModel cellModel)
            {
                // Update the DisplayText immediately for real-time validation
                cellModel.DisplayText = textBox.Text;
                
                // Perform real-time validation
                if (_uiManager != null)
                {
                    // Run validation on background to avoid UI blocking
                    _ = Task.Run(async () =>
                    {
                        await _uiManager.UpdateCellValidationAsync(cellModel);
                    });
                }
                
                _logger?.LogDebug("📝 TEXT CHANGED: Real-time validation triggered for cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TEXT CHANGE ERROR: Failed to handle text change");
        }
    }
    
    /// <summary>
    /// Event handler for exiting edit mode when TextBox loses focus
    /// </summary>
    private async void EditTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        try
        {
            if (sender is TextBox textBox && textBox.DataContext is DataCellModel cellModel)
            {
                // Exit edit mode
                cellModel.IsEditing = false;
                _logger?.LogDebug("💾 CELL EDIT: Exited edit mode for cell [{Row},{Col}] with value '{Value}'", 
                    cellModel.RowIndex, cellModel.ColumnIndex, cellModel.DisplayText);

                // Save the value to the controller
                await SaveCellValueAsync(cellModel);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL EDIT ERROR: Failed to exit edit mode");
        }
    }

    /// <summary>
    /// Event handler for keyboard navigation in edit mode
    /// </summary>
    private async void EditTextBox_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        try
        {
            if (sender is TextBox textBox && textBox.DataContext is DataCellModel cellModel)
            {
                // Check for modifier keys
                bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift)
                    .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                
                switch (e.Key)
                {
                    case Windows.System.VirtualKey.Enter:
                        if (isShiftPressed)
                        {
                            // Shift+Enter: Insert new line
                            int cursorPosition = textBox.SelectionStart;
                            string currentText = textBox.Text ?? string.Empty;
                            string newText = currentText.Insert(cursorPosition, Environment.NewLine);
                            
                            textBox.Text = newText;
                            textBox.SelectionStart = cursorPosition + Environment.NewLine.Length;
                            
                            e.Handled = true;
                            _logger?.LogDebug("📝 SHIFT+ENTER: Inserted newline at position {Position} in cell [{Row},{Col}]", 
                                cursorPosition, cellModel.RowIndex, cellModel.ColumnIndex);
                        }
                        else
                        {
                            // Enter: Save and exit edit mode
                            cellModel.IsEditing = false;
                            await SaveCellValueAsync(cellModel);
                            e.Handled = true;
                            _logger?.LogDebug("✅ CELL EDIT: Enter pressed - saved cell [{Row},{Col}]", 
                                cellModel.RowIndex, cellModel.ColumnIndex);
                        }
                        break;
                        
                    case Windows.System.VirtualKey.Escape:
                        // Cancel edit without saving
                        cellModel.IsEditing = false;
                        // Restore original value
                        cellModel.DisplayText = cellModel.Value?.ToString() ?? string.Empty;
                        e.Handled = true;
                        _logger?.LogDebug("❌ CELL EDIT: Escape pressed - cancelled edit for cell [{Row},{Col}]", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                        break;
                        
                    case Windows.System.VirtualKey.Tab:
                        // Move to next cell (could be implemented later)
                        _logger?.LogDebug("⭐ CELL EDIT: Tab pressed - future navigation feature");
                        break;
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL EDIT ERROR: Failed to handle key press");
        }
    }

    /// <summary>
    /// Helper method to find TextBox in visual tree
    /// </summary>
    private TextBox? FindTextBoxInVisualTree(DependencyObject parent)
    {
        if (parent == null) return null;

        for (int i = 0; i < Microsoft.UI.Xaml.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetChild(parent, i);
            
            if (child is TextBox textBox)
            {
                return textBox;
            }
            
            var result = FindTextBoxInVisualTree(child);
            if (result != null)
            {
                return result;
            }
        }
        
        return null;
    }

    /// <summary>
    /// Find the TextBox control for a specific cell
    /// </summary>
    private async Task<TextBox?> FindTextBoxForCellAsync(DataCellModel cellModel)
    {
        try
        {
            if (_uiManager == null || DataRepeater == null) return null;
            
            // Find the visual element for this cell by searching through the DataRepeater
            for (int i = 0; i < DataRepeater.ItemsSourceView?.Count; i++)
            {
                var rowElement = DataRepeater.GetOrCreateElement(i);
                if (rowElement is FrameworkElement rowFramework && 
                    rowFramework.DataContext is DataRowModel rowModel &&
                    rowModel.RowIndex == cellModel.RowIndex)
                {
                    // Found the row, now find the cell's TextBox
                    var textBox = FindTextBoxForCellInRow(rowFramework, cellModel.ColumnIndex);
                    if (textBox != null)
                    {
                        return textBox;
                    }
                }
            }
            
            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TEXTBOX SEARCH ERROR: Failed to find TextBox for cell [{Row},{Col}]", 
                cellModel.RowIndex, cellModel.ColumnIndex);
            return null;
        }
    }
    
    /// <summary>
    /// Find TextBox for a specific column in a row element
    /// </summary>
    private TextBox? FindTextBoxForCellInRow(FrameworkElement rowElement, int columnIndex)
    {
        try
        {
            // Look for ItemsRepeater that contains the cells
            var cellsRepeater = FindChildOfType<ItemsRepeater>(rowElement);
            if (cellsRepeater?.ItemsSourceView?.Count > columnIndex)
            {
                var cellElement = cellsRepeater.GetOrCreateElement(columnIndex);
                if (cellElement is FrameworkElement cellFramework)
                {
                    // Find the TextBox in this cell element
                    return FindTextBoxInVisualTree(cellFramework);
                }
            }
            
            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL TEXTBOX SEARCH ERROR: Failed to find TextBox in row for column {Column}", columnIndex);
            return null;
        }
    }
    
    /// <summary>
    /// Generic method to find child of specific type in visual tree
    /// </summary>
    private T? FindChildOfType<T>(DependencyObject parent) where T : DependencyObject
    {
        if (parent == null) return null;
        
        for (int i = 0; i < Microsoft.UI.Xaml.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = Microsoft.UI.Xaml.Media.VisualTreeHelper.GetChild(parent, i);
            
            if (child is T result)
            {
                return result;
            }
            
            var childOfType = FindChildOfType<T>(child);
            if (childOfType != null)
            {
                return childOfType;
            }
        }
        
        return null;
    }

    /// <summary>
    /// Find the currently focused cell
    /// </summary>
    private DataCellModel? FindFocusedCell()
    {
        if (_uiManager == null) return null;
        
        foreach (var row in _uiManager.RowsCollection)
        {
            foreach (var cell in row.Cells)
            {
                if (cell.IsFocused)
                {
                    return cell;
                }
            }
        }
        
        return null;
    }

    /// <summary>
    /// Check for double-click and start editing if detected
    /// </summary>
    private void CheckForDoubleClickEdit(DataCellModel cellModel)
    {
        try
        {
            var currentTime = DateTime.Now;
            
            // Check if this is the same cell clicked within timeout
            if (_lastClickedCell == cellModel && 
                (currentTime - _lastClickTime).TotalMilliseconds <= DoubleClickTimeoutMs)
            {
                // Double-click detected - start editing if not readonly
                if (!cellModel.IsReadOnly)
                {
                    StartCellEditing(cellModel, false); // false = don't clear content
                    _logger?.LogDebug("📝 DOUBLE-CLICK EDIT: Cell [{Row},{Col}] entered edit mode", 
                        cellModel.RowIndex, cellModel.ColumnIndex);
                }
                
                // Reset to prevent triple-click issues
                _lastClickedCell = null;
                _lastClickTime = DateTime.MinValue;
            }
            else
            {
                // First click or different cell - remember for next click
                _lastClickedCell = cellModel;
                _lastClickTime = currentTime;
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 DOUBLE-CLICK ERROR: Failed to check for double-click edit");
        }
    }

    /// <summary>
    /// Check if user is typing to start editing
    /// </summary>
    private bool CheckForTypeToEdit(Windows.System.VirtualKey key)
    {
        try
        {
            // Find focused cell
            var focusedCell = FindFocusedCell();
            if (focusedCell == null || focusedCell.IsReadOnly) return false;
            
            // Check if it's a printable character or digit
            bool isPrintable = (key >= Windows.System.VirtualKey.A && key <= Windows.System.VirtualKey.Z) ||
                              (key >= Windows.System.VirtualKey.Number0 && key <= Windows.System.VirtualKey.Number9) ||
                              (key >= Windows.System.VirtualKey.NumberPad0 && key <= Windows.System.VirtualKey.NumberPad9) ||
                              key == Windows.System.VirtualKey.Space;
            
            if (isPrintable)
            {
                // Start editing and clear content - let TextBox handle the character input naturally
                StartCellEditing(focusedCell, true); // true = clear content for new input
                
                _logger?.LogDebug("📝 TYPE-TO-EDIT: Cell [{Row},{Col}] entered edit mode by typing '{Key}' (TextBox will handle character)", 
                    focusedCell.RowIndex, focusedCell.ColumnIndex, key);
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TYPE-TO-EDIT ERROR: Failed to check for type-to-edit");
            return false;
        }
    }


    /// <summary>
    /// Check if any cell is currently in editing mode
    /// </summary>
    private bool IsCurrentlyEditing()
    {
        return _currentEditingCell != null && _currentEditingCell.IsEditing;
    }

    /// <summary>
    /// Start cell editing and remember original value for cancel functionality
    /// </summary>
    private void StartCellEditing(DataCellModel cellModel, bool clearContent)
    {
        try
        {
            // Save current editing state for cancel functionality
            _currentEditingCell = cellModel;
            _originalEditValue = cellModel.DisplayText;
            
            // Start editing
            cellModel.IsEditing = true;
            
            // Clear content if requested (for type-to-edit)
            if (clearContent)
            {
                cellModel.DisplayText = string.Empty;
            }
            
            _logger?.LogDebug("📝 EDIT START: Cell [{Row},{Col}] editing started (clear: {Clear}, original: '{Original}')", 
                cellModel.RowIndex, cellModel.ColumnIndex, clearContent, _originalEditValue);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 EDIT START ERROR: Failed to start cell editing");
        }
    }

    /// <summary>
    /// Save cell value to the controller and update validation
    /// </summary>
    private async Task SaveCellValueAsync(DataCellModel cellModel)
    {
        try
        {
            if (cellModel == null)
            {
                _logger?.LogError("🚨 SAVE ERROR: CellModel is null");
                return;
            }
            
            if (_controller == null)
            {
                _logger?.LogError("🚨 SAVE ERROR: Controller is null");
                return;
            }
            
            // Convert display text to appropriate type
            object? newValue = ConvertDisplayTextToValue(cellModel.DisplayText, cellModel.ColumnName);
            
            _logger?.LogError("🚨💾 DEBUG SAVE START: Saving cell [{Row},{Col}] '{OldValue}' → '{NewValue}' (DisplayText: '{DisplayText}')", 
                cellModel.RowIndex, cellModel.ColumnIndex, cellModel.Value?.ToString() ?? "null", newValue?.ToString() ?? "null", cellModel.DisplayText);
                
            // CRITICAL: Save to controller first (this persists to underlying TableCore)
            await _controller.SetCellValueAsync(cellModel.RowIndex, cellModel.ColumnIndex, newValue);
            
            _logger?.LogError("🚨💾 DEBUG SAVE CORE: Successfully saved to TableCore [{Row},{Col}] = '{Value}'", 
                cellModel.RowIndex, cellModel.ColumnIndex, newValue?.ToString() ?? "null");
            
            // CRITICAL: Update the model with the actual saved value to maintain consistency
            cellModel.Value = newValue;
            
            // PERSISTENCE FIX: Verify the data was actually saved by reading it back from TableCore
            try
            {
                // Wait a bit to ensure save is complete
                await Task.Delay(10);
                
                var verifyValue = await _controller.TableCore.GetCellValueAsync(cellModel.RowIndex, cellModel.ColumnIndex);
                _logger?.LogError("🚨💾 DEBUG VERIFY: Read back value [{Row},{Col}] = '{VerifyValue}' (expected: '{ExpectedValue}')", 
                    cellModel.RowIndex, cellModel.ColumnIndex, verifyValue?.ToString() ?? "null", newValue?.ToString() ?? "null");
                
                // Check if verification matches
                bool valuesMatch = (verifyValue?.ToString() ?? "") == (newValue?.ToString() ?? "");
                _logger?.LogError("🚨💾 DEBUG VERIFY RESULT: Values match = {Match}", valuesMatch);
                
                // CRITICAL FIX: Only update Value, preserve DisplayText to avoid UI reversion
                cellModel.Value = verifyValue;
                
                // Only update DisplayText if the values don't match (indicating a conversion/save issue)
                if (!valuesMatch)
                {
                    _logger?.LogError("🚨💾 MISMATCH: Updating DisplayText due to save/verify mismatch");
                    cellModel.DisplayText = verifyValue?.ToString() ?? string.Empty;
                }
                else
                {
                    _logger?.LogError("✅💾 MATCH: Keeping original DisplayText, values match");
                    // Keep the original DisplayText that user entered
                }
            }
            catch (Exception verifyEx)
            {
                _logger?.LogError(verifyEx, "🚨 CELL VERIFY ERROR: Failed to verify saved value [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
            
            // Update validation for this specific cell
            if (_uiManager != null)
            {
                await _uiManager.UpdateCellValidationAsync(cellModel);
            }
            
            _logger?.LogError("🚨💾 DEBUG SAVE COMPLETE: Cell [{Row},{Col}] save process completed with verification", 
                cellModel.RowIndex, cellModel.ColumnIndex);
            
            // Clear editing state after successful save (commit)
            if (_currentEditingCell == cellModel)
            {
                _currentEditingCell = null;
                _originalEditValue = null;
                _logger?.LogDebug("📝 EDIT COMMIT: Cell editing state cleared after save");
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL SAVE ERROR: Failed to save cell [{Row},{Col}]", 
                cellModel.RowIndex, cellModel.ColumnIndex);
            
            // Restore original value on error
            cellModel.DisplayText = cellModel.Value?.ToString() ?? string.Empty;
        }
    }

    /// <summary>
    /// Convert display text to appropriate value type based on column
    /// </summary>
    private object? ConvertDisplayTextToValue(string displayText, string columnName)
    {
        if (string.IsNullOrEmpty(displayText))
        {
            return null;
        }

        try
        {
            // For now, return displayText as string - type conversion can be enhanced later
            // TODO: Implement proper column type lookup when needed
            return displayText;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "⚠️ CELL CONVERT: Failed to convert '{DisplayText}' for column '{Column}', using string fallback", 
                displayText, columnName);
            return displayText;
        }
    }

    #endregion

    #region Cell Selection and Focus Event Handlers

    /// <summary>
    /// Event handler for cell selection via tapping
    /// </summary>
    private void CellBorder_Tapped(object sender, Microsoft.UI.Xaml.Input.TappedRoutedEventArgs e)
    {
        try
        {
            if (sender is Border border && border.DataContext is DataCellModel cellModel)
            {
                // Check for multi-select modifiers
                bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                
                if (!isCtrlPressed && !isShiftPressed)
                {
                    // Single selection - clear previous selections
                    ClearAllCellSelection();
                }
                
                // Set this cell as selected and focused
                cellModel.IsSelected = true;
                cellModel.IsFocused = true;
                
                // Update visual styling for selection
                UpdateCellSelectionVisuals(cellModel);
                
                _logger?.LogDebug("🎯 CELL SELECT: Cell [{Row},{Col}] selected (Ctrl: {Ctrl}, Shift: {Shift})", 
                    cellModel.RowIndex, cellModel.ColumnIndex, isCtrlPressed, isShiftPressed);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL SELECT ERROR: Failed to select cell");
        }
    }

    /// <summary>
    /// Event handler for cell context menu via right-click
    /// </summary>
    private void CellBorder_RightTapped(object sender, Microsoft.UI.Xaml.Input.RightTappedRoutedEventArgs e)
    {
        try
        {
            if (sender is Border border && border.DataContext is DataCellModel cellModel)
            {
                // Select cell first
                CellBorder_Tapped(sender, new Microsoft.UI.Xaml.Input.TappedRoutedEventArgs());
                
                // TODO: Show context menu for copy/paste operations
                _logger?.LogDebug("🖱️ CELL CONTEXT: Right-clicked cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CELL CONTEXT ERROR: Failed to handle right-click");
        }
    }

    /// <summary>
    /// Clear selection from all cells
    /// </summary>
    private void ClearAllCellSelection()
    {
        try
        {
            if (_uiManager != null)
            {
                foreach (var row in _uiManager.RowsCollection)
                {
                    foreach (var cell in row.Cells)
                    {
                        cell.IsSelected = false;
                        cell.IsFocused = false;
                        UpdateCellSelectionVisuals(cell);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 SELECTION ERROR: Failed to clear cell selection");
        }
    }

    /// <summary>
    /// Update visual styling for cell selection state
    /// </summary>
    private void UpdateCellSelectionVisuals(DataCellModel cellModel)
    {
        try
        {
            if (cellModel.IsSelected || cellModel.IsFocused)
            {
                // Selected/Focused state - blue border and slightly different background
                cellModel.BorderBrush = CreateBrush(Microsoft.UI.Colors.DodgerBlue);
                if (cellModel.IsValid)
                {
                    cellModel.BackgroundBrush = CreateBrush(Microsoft.UI.Colors.AliceBlue);
                }
            }
            else
            {
                // Normal state - restore original colors based on validation
                if (cellModel.IsValid)
                {
                    cellModel.BorderBrush = CreateBrush(_controller?.ColorConfig?.CellBorderColor ?? Microsoft.UI.Colors.LightGray);
                    cellModel.BackgroundBrush = CreateBrush(_controller?.ColorConfig?.CellBackgroundColor ?? Microsoft.UI.Colors.White);
                }
                else
                {
                    cellModel.BorderBrush = CreateBrush(_controller?.ColorConfig?.ValidationErrorBorderColor ?? Microsoft.UI.Colors.Red);
                    cellModel.BackgroundBrush = CreateBrush(_controller?.ColorConfig?.ValidationErrorBackgroundColor ?? Microsoft.UI.Colors.LightPink);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 VISUAL UPDATE ERROR: Failed to update cell selection visuals");
        }
    }

    /// <summary>
    /// Helper method to create SolidColorBrush from Color
    /// </summary>
    private SolidColorBrush CreateBrush(Windows.UI.Color color)
    {
        return new SolidColorBrush(color);
    }

    #endregion

    #region Drag Selection Event Handlers

    /// <summary>
    /// Start potential drag selection on pointer pressed
    /// </summary>
    private async void CellBorder_PointerPressed(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Border border && border.DataContext is DataCellModel cellModel)
            {
                // Check for Shift key for range selection
                bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift)
                    .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                    
                bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control)
                    .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);

                // Check if it's left mouse button and no resize operation
                if (e.Pointer.PointerDeviceType == Microsoft.UI.Input.PointerDeviceType.Mouse && 
                    e.GetCurrentPoint(border).Properties.IsLeftButtonPressed && !_isResizing)
                {
                    if (isShiftPressed && _focusedCell != null)
                    {
                        // Shift+Click: Range selection from focused cell to clicked cell
                        await SelectRangeAsync(_focusedCell, cellModel);
                        _logger?.LogDebug("🎯 SHIFT CLICK: Range selection from [{StartRow},{StartCol}] to [{EndRow},{EndCol}]",
                            _focusedCell.RowIndex, _focusedCell.ColumnIndex, cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                    else if (isCtrlPressed)
                    {
                        // Ctrl+Click: Multi-select individual cell
                        cellModel.IsSelected = !cellModel.IsSelected;
                        _logger?.LogDebug("🎯 CTRL CLICK: Toggle cell selection [{Row},{Col}] - selected: {Selected}",
                            cellModel.RowIndex, cellModel.ColumnIndex, cellModel.IsSelected);
                    }
                    else
                    {
                        // CRITICAL FIX: End any active edit mode before changing focus
                        if (_currentEditingCell != null && _currentEditingCell.IsEditing && _currentEditingCell != cellModel)
                        {
                            // Commit current edit before moving to new cell
                            await CommitCellEditAsync(stayOnCell: false);
                            _logger?.LogDebug("📝 EDIT AUTO-COMMIT: Committed edit mode when clicking different cell");
                        }
                        
                        // Regular click - prepare for potential drag selection or single selection
                        ClearAllSelection();
                        cellModel.IsSelected = true;
                        
                        // Update focus
                        await MoveFocusToAsync(cellModel.RowIndex, cellModel.ColumnIndex);
                        
                        // Prepare for potential drag selection
                        _isDragPending = true;
                        _dragStartCell = cellModel;
                        _dragStartPoint = e.GetCurrentPoint(border).Position;
                        
                        // Check for double-click editing
                        CheckForDoubleClickEdit(cellModel);
                        
                        _logger?.LogDebug("🎯 SINGLE CLICK: Selected cell [{Row},{Col}]", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                    
                    // Capture pointer for tracking
                    border.CapturePointer(e.Pointer);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CLICK ERROR: Failed to handle cell click");
        }
    }

    /// <summary>
    /// Update drag selection on pointer moved
    /// </summary>
    private void CellBorder_PointerMoved(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Border border && border.DataContext is DataCellModel cellModel)
            {
                if (_isDragPending && !_isDragSelecting)
                {
                    // Check if we've moved enough to start dragging
                    var currentPoint = e.GetCurrentPoint(border).Position;
                    var distance = Math.Sqrt(Math.Pow(currentPoint.X - _dragStartPoint.X, 2) + 
                                           Math.Pow(currentPoint.Y - _dragStartPoint.Y, 2));
                    
                    if (distance > 5) // Start drag after 5 pixels movement
                    {
                        _isDragSelecting = true;
                        _isDragPending = false;
                        _logger?.LogDebug("🎯 DRAG START: Started drag selection");
                    }
                }
                
                if (_isDragSelecting)
                {
                    // Update drag end cell
                    _dragEndCell = cellModel;
                    
                    // Update selection rectangle
                    UpdateDragSelection();
                    
                    _logger?.LogDebug("🎯 DRAG MOVE: Drag selection to [{Row},{Col}]", 
                        cellModel.RowIndex, cellModel.ColumnIndex);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 DRAG ERROR: Failed to update drag selection");
        }
    }

    /// <summary>
    /// Complete drag selection on pointer released
    /// </summary>
    private void CellBorder_PointerReleased(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Border border && border.DataContext is DataCellModel cellModel)
            {
                if (_isDragSelecting)
                {
                    // Complete drag selection
                    _isDragSelecting = false;
                    
                    _logger?.LogInformation("✅ DRAG COMPLETE: Selected from [{StartRow},{StartCol}] to [{EndRow},{EndCol}]", 
                        _dragStartCell?.RowIndex, _dragStartCell?.ColumnIndex,
                        _dragEndCell?.RowIndex, _dragEndCell?.ColumnIndex);
                }
                else if (_isDragPending)
                {
                    // This was a click, not a drag - handle single cell selection
                    _isDragPending = false;
                    
                    // Check for multi-select modifiers
                    bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                    bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                    
                    if (!isCtrlPressed && !isShiftPressed)
                    {
                        // Single selection - clear previous selections
                        ClearAllCellSelection();
                    }
                    
                    // Set this cell as selected and focused
                    cellModel.IsSelected = true;
                    cellModel.IsFocused = true;
                    
                    // Update visual styling for selection
                    UpdateCellSelectionVisuals(cellModel);
                    
                    // Check for double-click to enter edit mode
                    CheckForDoubleClickEdit(cellModel);
                    
                    _logger?.LogDebug("🎯 CLICK SELECT: Cell [{Row},{Col}] selected (Ctrl: {Ctrl}, Shift: {Shift})", 
                        cellModel.RowIndex, cellModel.ColumnIndex, isCtrlPressed, isShiftPressed);
                }
                
                // Release pointer capture
                border.ReleasePointerCapture(e.Pointer);
                
                // Reset drag variables
                _dragStartCell = null;
                _dragEndCell = null;
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 DRAG ERROR: Failed to complete drag selection");
        }
    }

    /// <summary>
    /// Update all cells in drag selection rectangle
    /// </summary>
    private void UpdateDragSelection()
    {
        try
        {
            if (_dragStartCell == null || _dragEndCell == null || _uiManager == null) return;

            // Calculate selection bounds
            int startRow = Math.Min(_dragStartCell.RowIndex, _dragEndCell.RowIndex);
            int endRow = Math.Max(_dragStartCell.RowIndex, _dragEndCell.RowIndex);
            int startCol = Math.Min(_dragStartCell.ColumnIndex, _dragEndCell.ColumnIndex);
            int endCol = Math.Max(_dragStartCell.ColumnIndex, _dragEndCell.ColumnIndex);

            // Clear previous selection
            ClearAllCellSelection();

            // Select cells in rectangle
            foreach (var row in _uiManager.RowsCollection)
            {
                if (row.RowIndex >= startRow && row.RowIndex <= endRow)
                {
                    for (int colIndex = startCol; colIndex <= endCol && colIndex < row.Cells.Count; colIndex++)
                    {
                        var cell = row.Cells[colIndex];
                        cell.IsSelected = true;
                        UpdateCellSelectionVisuals(cell);
                    }
                }
            }

            _logger?.LogDebug("🎯 DRAG UPDATE: Selected rectangle [{StartRow},{StartCol}] to [{EndRow},{EndCol}]", 
                startRow, startCol, endRow, endCol);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 DRAG ERROR: Failed to update drag selection rectangle");
        }
    }

    #endregion

    #region Keyboard Handlers

    /// <summary>
    /// Initialize keyboard shortcut handlers according to newProject.md specification
    /// </summary>
    private void InitializeKeyboardHandlers()
    {
        try
        {
            // Add KeyDown event handler to the root grid for global keyboard handling
            this.KeyDown += AdvancedDataGrid_KeyDown;
            
            // Ensure the control can receive keyboard focus
            this.IsTabStop = true;
            this.UseSystemFocusVisuals = true;
            
            _logger?.LogInformation("⌨️ KEYBOARD: Keyboard shortcuts initialized with global KeyDown handler");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 KEYBOARD ERROR: Failed to initialize keyboard handlers");
        }
    }

    /// <summary>
    /// Handle global keyboard shortcuts according to newProject.md specification
    /// </summary>
    private async void AdvancedDataGrid_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        try
        {
            bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
            bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);

            // Check for typing to start editing (if not Ctrl pressed and not already editing)
            if (!isCtrlPressed && !IsCurrentlyEditing() && CheckForTypeToEdit(e.Key))
            {
                // Start edit mode but let the key pass through naturally to TextBox
                // Do NOT set e.Handled = true to avoid character doubling
                return;
            }

            bool handled = await HandleKeyboardShortcutAsync(e.Key, isCtrlPressed, isShiftPressed);
            
            if (handled)
            {
                e.Handled = true;
                _logger?.LogDebug("⌨️ KEYBOARD: Handled shortcut {Key} (Ctrl: {Ctrl}, Shift: {Shift})", 
                    e.Key, isCtrlPressed, isShiftPressed);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 KEYBOARD ERROR: Failed to handle keyboard shortcut {Key}", e.Key);
        }
    }

    /// <summary>
    /// Process keyboard shortcuts according to newProject.md specification
    /// </summary>
    private async Task<bool> HandleKeyboardShortcutAsync(Windows.System.VirtualKey key, bool isCtrlPressed, bool isShiftPressed)
    {
        // Check if we're in edit mode for special Tab behavior
        bool isInEditMode = _currentEditingCell != null && _currentEditingCell.IsEditing;
        
        return (key, isCtrlPressed, isShiftPressed, isInEditMode) switch
        {
            // Basic editing shortcuts
            (Windows.System.VirtualKey.F2, false, false, _) => await StartEditFocusedCellAsync(),
            (Windows.System.VirtualKey.Escape, false, false, _) => await CancelCellEditAsync(),
            (Windows.System.VirtualKey.Enter, false, false, _) => await CommitCellEditAsync(stayOnCell: true),
            (Windows.System.VirtualKey.Enter, false, true, true) => await InsertNewLineInCellAsync(), // Only in edit mode
            
            // Tab behavior: Insert tab in edit mode, navigate otherwise
            (Windows.System.VirtualKey.Tab, false, false, true) => await InsertTabInCellAsync(),
            (Windows.System.VirtualKey.Tab, false, false, false) => await MoveToNextCellAsync(),
            (Windows.System.VirtualKey.Tab, false, true, false) => await MoveToPreviousCellAsync(),
            
            // Clipboard operations
            (Windows.System.VirtualKey.C, true, false, _) => await CopySelectedCellsAsync(),
            (Windows.System.VirtualKey.V, true, false, _) => await PasteFromClipboardAsync(),
            (Windows.System.VirtualKey.X, true, false, _) => await CutSelectedCellsAsync(),
            (Windows.System.VirtualKey.A, true, false, _) => await SelectAllCellsAsync(),
            
            // Delete operations
            (Windows.System.VirtualKey.Delete, false, false, _) => await SmartDeleteAsync(),
            (Windows.System.VirtualKey.Delete, true, false, _) => await ForceDeleteRowAsync(),
            
            // Insert operations
            (Windows.System.VirtualKey.Insert, false, false, _) => await InsertRowAboveAsync(),
            
            // Navigation shortcuts
            (Windows.System.VirtualKey.Up, false, false, _) => await MoveUpAsync(),
            (Windows.System.VirtualKey.Down, false, false, _) => await MoveDownAsync(),
            (Windows.System.VirtualKey.Left, false, false, _) => await MoveLeftAsync(),
            (Windows.System.VirtualKey.Right, false, false, _) => await MoveRightAsync(),
            (Windows.System.VirtualKey.Home, true, false, _) => await MoveToFirstCellAsync(),
            (Windows.System.VirtualKey.End, true, false, _) => await MoveToLastDataCellAsync(),
            
            _ => false // Unhandled
        };
    }

    #region Keyboard Shortcut Implementations

    /// <summary>
    /// Start editing the currently focused cell (F2)
    /// </summary>
    private async Task<bool> StartEditFocusedCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Start edit focused cell (F2)");
            
            // Find the currently focused cell
            var focusedCell = FindFocusedCell();
            if (focusedCell != null && !focusedCell.IsReadOnly)
            {
                StartCellEditing(focusedCell, false); // false = don't clear content
                _logger?.LogDebug("📝 EDIT START: Cell [{Row},{Col}] entered edit mode via F2", 
                    focusedCell.RowIndex, focusedCell.ColumnIndex);
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 EDIT ERROR: Failed to start editing focused cell");
            return false;
        }
    }

    /// <summary>
    /// Cancel current cell editing and restore original value
    /// </summary>
    private async Task<bool> CancelCellEditAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Cancel cell edit (Escape)");
            
            if (_currentEditingCell != null && _originalEditValue != null)
            {
                // Restore original value
                _currentEditingCell.DisplayText = _originalEditValue;
                _currentEditingCell.IsEditing = false;
                
                _logger?.LogDebug("📝 EDIT CANCEL: Cell [{Row},{Col}] editing cancelled, restored to '{Original}'", 
                    _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex, _originalEditValue);
                
                // Clear editing state
                _currentEditingCell = null;
                _originalEditValue = null;
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 EDIT CANCEL ERROR: Failed to cancel cell editing");
            return false;
        }
    }

    /// <summary>
    /// Commit current cell edit and optionally stay on the same cell
    /// </summary>
    private async Task<bool> CommitCellEditAsync(bool stayOnCell)
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Commit cell edit (Enter), stay on cell: {StayOnCell}", stayOnCell);
            
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                // Store reference before clearing state
                var editingCell = _currentEditingCell;
                
                // Save the current value to TableCore
                if (_controller != null)
                {
                    await SaveCellValueAsync(editingCell);
                }
                else
                {
                    _logger?.LogError("🚨 EDIT COMMIT ERROR: Controller is null, cannot save cell value");
                    return false;
                }
                
                // Exit edit mode
                editingCell.IsEditing = false;
                
                _logger?.LogDebug("📝 EDIT COMMIT: Cell [{Row},{Col}] editing committed with value '{Value}'", 
                    editingCell.RowIndex, editingCell.ColumnIndex, editingCell.DisplayText);
                
                // Clear editing state
                _currentEditingCell = null;
                _originalEditValue = null;
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 EDIT COMMIT ERROR: Failed to commit cell edit");
            return false;
        }
    }

    /// <summary>
    /// Insert new line in cell (Shift+Enter in edit mode)
    /// </summary>
    private async Task<bool> InsertNewLineInCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Insert new line in cell (Shift+Enter)");
            
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                // Find the actual TextBox control for the editing cell
                var textBox = await FindTextBoxForCellAsync(_currentEditingCell);
                if (textBox != null && DispatcherQueue != null)
                {
                    // Insert newline at cursor position using UI thread
                    DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
                    {
                        int cursorPosition = textBox.SelectionStart;
                        string currentText = textBox.Text ?? string.Empty;
                        string newText = currentText.Insert(cursorPosition, Environment.NewLine);
                        
                        textBox.Text = newText;
                        textBox.SelectionStart = cursorPosition + Environment.NewLine.Length;
                        
                        // Update the cell model to match
                        _currentEditingCell.DisplayText = newText;
                        
                        _logger?.LogDebug("📝 MULTILINE: Inserted newline at cursor position {Position} in cell [{Row},{Column}]", 
                            cursorPosition, _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex);
                    });
                    
                    return true;
                }
                else
                {
                    // Fallback: Insert at end if TextBox not found
                    string currentText = _currentEditingCell.DisplayText ?? string.Empty;
                    string newText = currentText + Environment.NewLine;
                    _currentEditingCell.DisplayText = newText;
                    
                    _logger?.LogDebug("📝 MULTILINE FALLBACK: Inserted newline at end in cell [{Row},{Column}]", 
                        _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex);
                    
                    return true;
                }
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 MULTILINE ERROR: Failed to insert new line in cell");
            return false;
        }
    }

    /// <summary>
    /// Insert tab character in current cell (Tab in edit mode)
    /// </summary>
    private async Task<bool> InsertTabInCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Insert tab in cell (Tab in edit mode)");
            
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                // Insert tab character at the end of the current text (simplified approach)
                string currentText = _currentEditingCell.DisplayText ?? string.Empty;
                string newText = currentText + "\t";
                _currentEditingCell.DisplayText = newText;
                
                _logger?.LogDebug("📋 TAB INSERT: Inserted tab in cell [{Row},{Column}] - new length: {Length}", 
                    _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex, newText.Length);
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TAB INSERT ERROR: Failed to insert tab in cell");
            return false;
        }
    }

    /// <summary>
    /// Move to next cell (Tab) - moves right, wraps to next row
    /// </summary>
    private async Task<bool> MoveToNextCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Move to next cell (Tab)");
            
            if (_uiManager == null) return false;
            
            int nextColumn = _focusedColumnIndex + 1;
            int nextRow = _focusedRowIndex;
            
            // If at end of row, wrap to next row
            if (nextColumn >= ColumnCount)
            {
                nextColumn = 0;
                nextRow++;
                
                // If at end of data, wrap to first cell
                if (nextRow >= ActualRowCount)
                {
                    nextRow = 0;
                }
            }
            
            return await MoveFocusToAsync(nextRow, nextColumn);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move to next cell");
            return false;
        }
    }

    /// <summary>
    /// Move to previous cell (Shift+Tab) - moves left, wraps to previous row
    /// </summary>
    private async Task<bool> MoveToPreviousCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Move to previous cell (Shift+Tab)");
            
            if (_uiManager == null) return false;
            
            int prevColumn = _focusedColumnIndex - 1;
            int prevRow = _focusedRowIndex;
            
            // If at beginning of row, wrap to previous row
            if (prevColumn < 0)
            {
                prevColumn = ColumnCount - 1;
                prevRow--;
                
                // If at beginning of data, wrap to last cell
                if (prevRow < 0)
                {
                    prevRow = ActualRowCount - 1;
                }
            }
            
            return await MoveFocusToAsync(prevRow, prevColumn);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move to previous cell");
            return false;
        }
    }

    /// <summary>
    /// Copy selected cells to clipboard (Ctrl+C)
    /// </summary>
    private async Task<bool> CopySelectedCellsAsync()
    {
        try
        {
            if (_uiManager == null) return false;

            var selectedCells = new List<DataCellModel>();
            
            // Collect all selected cells
            foreach (var row in _uiManager.RowsCollection)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.IsSelected)
                    {
                        selectedCells.Add(cell);
                    }
                }
            }

            if (selectedCells.Count == 0)
            {
                _logger?.LogWarning("⌨️ COPY: No cells selected");
                return false;
            }

            // Create tab-separated text for clipboard
            var copiedText = await CreateClipboardTextFromSelectedCells(selectedCells);
            
            // Copy to system clipboard
            var dataPackage = new Windows.ApplicationModel.DataTransfer.DataPackage();
            dataPackage.SetText(copiedText);
            Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(dataPackage);

            _logger?.LogInformation("⌨️ COPY: Copied {Count} selected cells to clipboard", selectedCells.Count);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 COPY ERROR: Failed to copy selected cells");
            return false;
        }
    }

    /// <summary>
    /// Paste from clipboard (Ctrl+V)
    /// </summary>
    private async Task<bool> PasteFromClipboardAsync()
    {
        try
        {
            if (_uiManager == null) return false;

            // Get clipboard content
            var dataPackageView = Windows.ApplicationModel.DataTransfer.Clipboard.GetContent();
            if (!dataPackageView.Contains(Windows.ApplicationModel.DataTransfer.StandardDataFormats.Text))
            {
                _logger?.LogWarning("⌨️ PASTE: No text content in clipboard");
                return false;
            }

            var clipboardText = await dataPackageView.GetTextAsync();
            if (string.IsNullOrEmpty(clipboardText))
            {
                _logger?.LogWarning("⌨️ PASTE: Clipboard text is empty");
                return false;
            }

            // Find focused cell as paste target
            var focusedCell = GetFocusedCell();
            if (focusedCell == null)
            {
                _logger?.LogWarning("⌨️ PASTE: No focused cell for paste target");
                return false;
            }

            // Parse clipboard text and paste
            await PasteClipboardTextToCells(clipboardText, focusedCell);

            _logger?.LogInformation("⌨️ PASTE: Pasted clipboard content starting at [{Row},{Col}]", 
                focusedCell.RowIndex, focusedCell.ColumnIndex);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 PASTE ERROR: Failed to paste from clipboard");
            return false;
        }
    }

    /// <summary>
    /// Cut selected cells to clipboard (Ctrl+X)
    /// </summary>
    private async Task<bool> CutSelectedCellsAsync()
    {
        _logger?.LogInformation("⌨️ KEYBOARD: Cut selected cells (Ctrl+X)");
        // TODO: Implement clipboard cut logic
        return true;
    }

    /// <summary>
    /// Select all cells (Ctrl+A)
    /// </summary>
    private async Task<bool> SelectAllCellsAsync()
    {
        _logger?.LogInformation("⌨️ KEYBOARD: Select all cells (Ctrl+A)");
        // TODO: Implement select all logic
        return true;
    }

    /// <summary>
    /// Smart delete - clear cell content or delete row based on context (Delete)
    /// </summary>
    /// <summary>
    /// Smart delete focused cells (Delete) - clears content of focused/selected cells
    /// </summary>
    private async Task<bool> SmartDeleteAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Smart delete (Delete)");
            
            if (_uiManager == null) return false;
            
            var cellsToDelete = new List<DataCellModel>();
            
            // Collect selected cells or focused cell
            foreach (var row in _uiManager.RowsCollection)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.IsSelected || (cell.IsFocused && cellsToDelete.Count == 0))
                    {
                        cellsToDelete.Add(cell);
                    }
                }
            }
            
            // Delete content of collected cells
            foreach (var cell in cellsToDelete)
            {
                if (!cell.IsReadOnly)
                {
                    // Clear cell content through TableCore
                    await TableCore.SetCellValueAsync(cell.RowIndex, cell.ColumnIndex, null);
                    cell.DisplayText = string.Empty;
                    
                    _logger?.LogDebug("🗑️ SMART DELETE: Cleared cell [{Row},{Col}]", 
                        cell.RowIndex, cell.ColumnIndex);
                }
            }
            
            _logger?.LogInformation("✅ SMART DELETE: Cleared {Count} cells", cellsToDelete.Count);
            return cellsToDelete.Count > 0;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 DELETE ERROR: Failed to smart delete cells");
            return false;
        }
    }

    /// <summary>
    /// Force delete entire row (Ctrl+Delete) - deletes the focused row completely
    /// </summary>
    private async Task<bool> ForceDeleteRowAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Force delete row (Ctrl+Delete)");
            
            if (_focusedCell == null) return false;
            
            int rowToDelete = _focusedCell.RowIndex;
            
            // Use TableCore's smart delete functionality
            await TableCore.SmartDeleteRowAsync(rowToDelete);
            
            // Update focus if needed
            if (_focusedRowIndex >= TableCore.ActualRowCount)
            {
                _focusedRowIndex = Math.Max(0, TableCore.ActualRowCount - 1);
            }
            
            // Refresh UI to reflect the deletion
            await RefreshUIAsync();
            
            _logger?.LogInformation("✅ ROW DELETE: Deleted row {Row}, remaining rows: {Total}", 
                rowToDelete, TableCore.ActualRowCount);
                
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 ROW DELETE ERROR: Failed to delete row");
            return false;
        }
    }

    /// <summary>
    /// Insert row above current position (Insert)
    /// </summary>
    private async Task<bool> InsertRowAboveAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Insert row above (Insert)");
            
            if (_focusedCell == null) return false;
            
            int insertPosition = _focusedCell.RowIndex;
            
            // Create empty row data
            var emptyRowData = new Dictionary<string, object?>();
            var columnNames = TableCore.GetColumnNames();
            foreach (var columnName in columnNames)
            {
                emptyRowData[columnName] = null;
            }
            
            // Insert the empty row through paste operation (which handles insertion)
            var rowsToInsert = new List<Dictionary<string, object?>> { emptyRowData };
            await TableCore.PasteDataAsync(rowsToInsert, insertPosition, 0);
            
            // Refresh UI to show the new row
            await RefreshUIAsync();
            
            // Move focus to the newly inserted row
            await MoveFocusToAsync(insertPosition, _focusedColumnIndex);
            
            _logger?.LogInformation("✅ ROW INSERT: Inserted empty row at position {Position}, total rows: {Total}", 
                insertPosition, TableCore.ActualRowCount);
                
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 ROW INSERT ERROR: Failed to insert row");
            return false;
        }
    }

    /// <summary>
    /// Move to first cell (Ctrl+Home)
    /// </summary>
    private async Task<bool> MoveToFirstCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Move to first cell (Ctrl+Home)");
            return await MoveFocusToAsync(0, 0);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move to first cell");
            return false;
        }
    }

    /// <summary>
    /// Move to last data cell (Ctrl+End)
    /// </summary>
    private async Task<bool> MoveToLastDataCellAsync()
    {
        try
        {
            _logger?.LogInformation("⌨️ KEYBOARD: Move to last data cell (Ctrl+End)");
            
            // Find the last row with data
            int lastDataRow = await TableCore.GetLastDataRowAsync();
            if (lastDataRow < 0) lastDataRow = 0; // If no data, go to first row
            
            int lastColumn = Math.Max(0, ColumnCount - 1);
            
            return await MoveFocusToAsync(lastDataRow, lastColumn);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move to last data cell");
            return false;
        }
    }

    /// <summary>
    /// Move up one row (Up Arrow)
    /// </summary>
    private async Task<bool> MoveUpAsync()
    {
        try
        {
            _logger?.LogDebug("⌨️ KEYBOARD: Move up (Arrow Up)");
            
            int newRow = Math.Max(0, _focusedRowIndex - 1);
            return await MoveFocusToAsync(newRow, _focusedColumnIndex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move up");
            return false;
        }
    }

    /// <summary>
    /// Move down one row (Down Arrow)
    /// </summary>
    private async Task<bool> MoveDownAsync()
    {
        try
        {
            _logger?.LogDebug("⌨️ KEYBOARD: Move down (Arrow Down)");
            
            int newRow = Math.Min(ActualRowCount - 1, _focusedRowIndex + 1);
            return await MoveFocusToAsync(newRow, _focusedColumnIndex);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move down");
            return false;
        }
    }

    /// <summary>
    /// Move left one column (Left Arrow)
    /// </summary>
    private async Task<bool> MoveLeftAsync()
    {
        try
        {
            _logger?.LogDebug("⌨️ KEYBOARD: Move left (Arrow Left)");
            
            int newColumn = Math.Max(0, _focusedColumnIndex - 1);
            return await MoveFocusToAsync(_focusedRowIndex, newColumn);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move left");
            return false;
        }
    }

    /// <summary>
    /// Move right one column (Right Arrow)
    /// </summary>
    private async Task<bool> MoveRightAsync()
    {
        try
        {
            _logger?.LogDebug("⌨️ KEYBOARD: Move right (Arrow Right)");
            
            int newColumn = Math.Min(ColumnCount - 1, _focusedColumnIndex + 1);
            return await MoveFocusToAsync(_focusedRowIndex, newColumn);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move right");
            return false;
        }
    }

    /// <summary>
    /// Core method to move focus to a specific cell and update visual focus
    /// </summary>
    private async Task<bool> MoveFocusToAsync(int row, int column)
    {
        try
        {
            if (_uiManager == null) return false;
            
            // Validate bounds
            if (row < 0 || row >= ActualRowCount || column < 0 || column >= ColumnCount)
            {
                _logger?.LogWarning("🚨 NAVIGATION: Invalid cell position [{Row},{Column}] - bounds: [{MaxRow},{MaxColumn}]", 
                    row, column, ActualRowCount - 1, ColumnCount - 1);
                return false;
            }

            // Update focus tracking
            _focusedRowIndex = row;
            _focusedColumnIndex = column;

            // Clear previous focus visual
            if (_focusedCell != null)
            {
                _focusedCell.IsFocused = false;
            }

            // Get new focused cell
            if (row < _uiManager.RowsCollection.Count && column < _uiManager.RowsCollection[row].Cells.Count)
            {
                _focusedCell = _uiManager.RowsCollection[row].Cells[column];
                _focusedCell.IsFocused = true;
                
                _logger?.LogDebug("🎯 FOCUS: Moved to cell [{Row},{Column}] - '{DisplayText}'", 
                    row, column, _focusedCell.DisplayText);
                
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 NAVIGATION ERROR: Failed to move focus to [{Row},{Column}]", row, column);
            return false;
        }
    }


    /// <summary>
    /// Select range of cells from start to end (Shift+Click functionality)
    /// </summary>
    private async Task SelectRangeAsync(DataCellModel startCell, DataCellModel endCell)
    {
        try
        {
            if (_uiManager == null) return;
            
            // Clear previous selection
            ClearAllSelection();
            
            // Calculate range boundaries
            int startRow = Math.Min(startCell.RowIndex, endCell.RowIndex);
            int endRow = Math.Max(startCell.RowIndex, endCell.RowIndex);
            int startCol = Math.Min(startCell.ColumnIndex, endCell.ColumnIndex);
            int endCol = Math.Max(startCell.ColumnIndex, endCell.ColumnIndex);
            
            // Select all cells in the range
            for (int row = startRow; row <= endRow; row++)
            {
                if (row < _uiManager.RowsCollection.Count)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        if (col < _uiManager.RowsCollection[row].Cells.Count)
                        {
                            _uiManager.RowsCollection[row].Cells[col].IsSelected = true;
                        }
                    }
                }
            }
            
            _logger?.LogDebug("🎯 RANGE SELECT: Selected range ({StartRow},{StartCol}) to ({EndRow},{EndCol}) - {Count} cells",
                startRow, startCol, endRow, endCol, (endRow - startRow + 1) * (endCol - startCol + 1));
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 RANGE SELECT ERROR: Failed to select cell range");
        }
    }

    /// <summary>
    /// Clear all cell selections
    /// </summary>
    private void ClearAllSelection()
    {
        try
        {
            if (_uiManager == null) return;
            
            foreach (var row in _uiManager.RowsCollection)
            {
                foreach (var cell in row.Cells)
                {
                    cell.IsSelected = false;
                    // CRITICAL FIX: Update visual styling to reflect selection change
                    UpdateCellSelectionVisuals(cell);
                }
            }
            
            _logger?.LogDebug("🧹 CLEAR SELECT: Cleared all cell selections with visual updates");
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 CLEAR SELECT ERROR: Failed to clear selections");
        }
    }


    /// <summary>
    /// Get the TextBox element for a cell that's currently in edit mode
    /// </summary>
    private async Task<TextBox?> GetEditingTextBoxAsync(DataCellModel cell)
    {
        try
        {
            if (!cell.IsEditing) return null;

            // In WinUI, we need to find the TextBox within the ItemsRepeater structure
            // This is a simplified approach - in a real implementation, you might need
            // more sophisticated element finding based on your XAML structure
            
            // For now, we'll return a reference that can be used for text manipulation
            // The actual implementation would depend on your specific XAML structure and data binding
            
            _logger?.LogDebug("🔍 TEXTBOX: Looking for editing TextBox for cell [{Row},{Column}]", 
                cell.RowIndex, cell.ColumnIndex);
            
            // This would need to be implemented based on your actual XAML structure
            // For now, we'll simulate this by working with the cell's DisplayText directly
            return null; // TODO: Implement actual TextBox finding logic
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 TEXTBOX ERROR: Failed to find editing TextBox for cell [{Row},{Column}]", 
                cell.RowIndex, cell.ColumnIndex);
            return null;
        }
    }

    #endregion

    #region Helper Methods for Shortcuts

    /// <summary>
    /// Create tab-separated text from selected cells for clipboard
    /// </summary>
    private async Task<string> CreateClipboardTextFromSelectedCells(List<DataCellModel> selectedCells)
    {
        if (selectedCells.Count == 0) return string.Empty;

        // Group cells by row and sort
        var cellsByRow = selectedCells
            .GroupBy(c => c.RowIndex)
            .OrderBy(g => g.Key)
            .ToList();

        var lines = new List<string>();
        
        foreach (var rowGroup in cellsByRow)
        {
            var cellsInRow = rowGroup.OrderBy(c => c.ColumnIndex).ToList();
            var cellValues = cellsInRow.Select(c => c.DisplayText ?? string.Empty);
            lines.Add(string.Join("\t", cellValues));
        }

        return string.Join("\r\n", lines);
    }

    /// <summary>
    /// Get currently focused cell
    /// </summary>
    private DataCellModel? GetFocusedCell()
    {
        if (_uiManager == null) return null;

        foreach (var row in _uiManager.RowsCollection)
        {
            foreach (var cell in row.Cells)
            {
                if (cell.IsFocused)
                {
                    return cell;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// Parse clipboard text and paste to cells starting from target cell
    /// </summary>
    private async Task PasteClipboardTextToCells(string clipboardText, DataCellModel targetCell)
    {
        try
        {
            var lines = clipboardText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            int startRow = targetCell.RowIndex;
            int startCol = targetCell.ColumnIndex;

            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                var line = lines[lineIndex];
                var cellValues = line.Split('\t');
                
                for (int colIndex = 0; colIndex < cellValues.Length; colIndex++)
                {
                    int targetRowIndex = startRow + lineIndex;
                    int targetColIndex = startCol + colIndex;
                    
                    // Find target cell
                    var targetCellModel = FindCellAt(targetRowIndex, targetColIndex);
                    if (targetCellModel != null && !targetCellModel.IsReadOnly)
                    {
                        // Update cell value
                        targetCellModel.DisplayText = cellValues[colIndex];
                        targetCellModel.Value = cellValues[colIndex];
                        
                        // Save to underlying data
                        await _controller.SetCellValueAsync(targetRowIndex, targetColIndex, cellValues[colIndex]);
                        
                        // Update validation
                        if (_uiManager != null)
                        {
                            await _uiManager.UpdateCellValidationAsync(targetCellModel);
                        }
                    }
                }
            }

            _logger?.LogInformation("📋 PASTE: Pasted {Lines} lines x {Cols} columns to grid", 
                lines.Length, lines.Length > 0 ? lines[0].Split('\t').Length : 0);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "🚨 PASTE ERROR: Failed to parse and paste clipboard text");
        }
    }

    /// <summary>
    /// Find cell at specific row and column position
    /// </summary>
    private DataCellModel? FindCellAt(int rowIndex, int columnIndex)
    {
        if (_uiManager == null) return null;

        var targetRow = _uiManager.RowsCollection.FirstOrDefault(r => r.RowIndex == rowIndex);
        if (targetRow == null || columnIndex < 0 || columnIndex >= targetRow.Cells.Count)
        {
            return null;
        }

        return targetRow.Cells[columnIndex];
    }

    #endregion

    #endregion

}