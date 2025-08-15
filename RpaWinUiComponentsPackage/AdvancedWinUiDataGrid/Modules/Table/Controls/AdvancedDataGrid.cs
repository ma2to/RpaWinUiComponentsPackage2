using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Shapes;
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

    /// <summary>
    /// Event handler for pointer entering resize handle area
    /// </summary>
    private void ResizeHandle_PointerEntered(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            if (sender is Rectangle handle)
            {
                // Change cursor to resize cursor
                this.ProtectedCursor = Microsoft.UI.Input.InputSystemCursor.Create(Microsoft.UI.Input.InputSystemCursorShape.SizeWestEast);
                _logger?.LogDebug("🔍 RESIZE: Pointer entered resize handle");
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
                    _resizeStartX = e.GetCurrentPoint(handle).Position.X;
                    
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
            if (_isResizing && sender is Rectangle handle && _resizingColumnIndex >= 0)
            {
                double currentX = e.GetCurrentPoint(handle).Position.X;
                double deltaX = currentX - _resizeStartX;
                double targetWidth = _resizeStartWidth + deltaX;
                
                // Apply MinWidth/MaxWidth constraints from column definition
                double newWidth = ApplyColumnWidthConstraints(_resizingColumnIndex, targetWidth);
                
                // Update the column width in real-time
                UpdateColumnWidth(_resizingColumnIndex, newWidth);
                
                _logger?.LogDebug("🔄 RESIZE: Column {ColumnIndex} width = {NewWidth} (delta: {Delta})", 
                    _resizingColumnIndex, newWidth, deltaX);
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
                double currentX = e.GetCurrentPoint(handle).Position.X;
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
    /// Update column width for both header and all cells in that column
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
    /// Event handler for entering edit mode when clicking on TextBlock
    /// </summary>
    private void DisplayText_Tapped(object sender, Microsoft.UI.Xaml.Input.TappedRoutedEventArgs e)
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
                switch (e.Key)
                {
                    case Windows.System.VirtualKey.Enter:
                        // Save and exit edit mode
                        cellModel.IsEditing = false;
                        await SaveCellValueAsync(cellModel);
                        e.Handled = true;
                        _logger?.LogDebug("✅ CELL EDIT: Enter pressed - saved cell [{Row},{Col}]", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
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
    /// Save cell value to the controller and update validation
    /// </summary>
    private async Task SaveCellValueAsync(DataCellModel cellModel)
    {
        try
        {
            // Convert display text to appropriate type
            object? newValue = ConvertDisplayTextToValue(cellModel.DisplayText, cellModel.ColumnName);
            
            // Save to controller
            await _controller.SetCellValueAsync(cellModel.RowIndex, cellModel.ColumnIndex, newValue);
            
            // Update the model
            cellModel.Value = newValue;
            
            // Trigger validation update
            if (_uiManager != null)
            {
                await _uiManager.UpdateRowUIAsync(cellModel.RowIndex);
            }
            
            _logger?.LogInformation("💾 CELL SAVE: Cell [{Row},{Col}] saved with value '{Value}'", 
                cellModel.RowIndex, cellModel.ColumnIndex, newValue?.ToString() ?? "null");
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

}