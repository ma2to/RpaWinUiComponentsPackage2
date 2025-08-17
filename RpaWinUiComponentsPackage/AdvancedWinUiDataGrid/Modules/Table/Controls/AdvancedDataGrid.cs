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
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

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
    
    // Note: Removed double-click detection fields as we now use direct second-click logic
    
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

            _logger?.Info("🎨 UI WRAPPER: AdvancedDataGrid UI layer initialized successfully");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI WRAPPER ERROR: AdvancedDataGrid UI initialization failed");
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
                                _logger?.Info("🎨 UI BINDING: HeaderRepeater bound to HeadersCollection on UI thread");
                            }
                            
                            if (DataRepeater != null)
                            {
                                DataRepeater.ItemsSource = _uiManager.RowsCollection;
                                _logger?.Info("🎨 UI BINDING: DataRepeater bound to RowsCollection on UI thread");
                            }
                            bindingSuccess = true;
                        }
                        catch (Exception bindEx)
                        {
                            _logger?.Error(bindEx, "🚨 UI BINDING ERROR: Failed to bind ItemsSource on UI thread");
                        }
                    });
                    
                    // Wait a moment for UI thread operation to complete
                    await Task.Delay(100);
                    _logger?.Info("✅ UI BINDING: Dispatcher binding request submitted, Success: {Success}", bindingSuccess);
                }
                else
                {
                    // Fallback to direct binding if no dispatcher available
                    _logger?.Warning("⚠️ UI BINDING: No DispatcherQueue available, using direct binding");
                    
                    if (HeaderRepeater != null)
                    {
                        HeaderRepeater.ItemsSource = _uiManager.HeadersCollection;
                        _logger?.Info("🎨 UI BINDING: HeaderRepeater bound to HeadersCollection (direct)");
                    }
                    
                    if (DataRepeater != null)
                    {
                        DataRepeater.ItemsSource = _uiManager.RowsCollection;
                        _logger?.Info("🎨 UI BINDING: DataRepeater bound to RowsCollection (direct)");
                    }
                }
            }

            _logger?.Info("✅ UI LAYER: Initialization completed");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI LAYER ERROR: UI layer initialization failed");
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
            
            // DATA RETENTION FIX: Commit any active edit operations before refresh
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                _logger?.Info("💾 DATA RETENTION: Auto-committing active edit before refresh");
                await CommitCellEditAsync(stayOnCell: false);
            }
            
            await ShowLoadingAsync(true);
            await RenderAllCellsAsync();
            
            // VALIDATION ALERTS FIX: Auto-update validation UI during refresh to ensure ValidationAlerts column is populated
            _logger?.Info("🔍 UI UPDATE: Auto-updating validation during refresh to populate ValidationAlerts");
            await UpdateValidationVisualsAsync();
            
            await ShowLoadingAsync(false);

            _logger?.Info("✅ UI UPDATE: Full UI refresh completed with data retention");
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
                
                _logger?.Info("🎨 UI LAYOUT: Updated with StackLayout to preserve column widths, row height: {Height}px", Math.Ceiling(currentRowHeight));
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

            _logger?.Info("🎨 XAML ELEMENTS: Updated programatically from color config (including templates)");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 XAML ERROR: UpdateXAMLProperties failed");
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
            
            _logger?.Info("🎨 CELL TEMPLATE: Colors will be applied during runtime rendering (WinUI3 compatible)");
            
            // Note: Actual cell coloring happens in ApplyColorConfiguration() and ZebraRowColorManager
            // Templates stay generic, colors are applied programmatically to rendered elements
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 TEMPLATE ERROR: UpdateCellDataTemplate failed");
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
            
            _logger?.Info("🎨 HEADER TEMPLATE: Colors will be applied during runtime rendering (WinUI3 compatible)");
            
            // Note: Actual header coloring happens in ApplyColorConfiguration() and ZebraRowColorManager
            // Templates stay generic, colors are applied programmatically to rendered elements
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 TEMPLATE ERROR: UpdateHeaderDataTemplate failed");
        }
    }

    /// <summary>
    /// Initialize UI event handlers - CRITICAL: Column resize & cell interaction
    /// </summary>
    private void InitializeUIEventHandlers()
    {
        try
        {
            _logger?.Info("🎨 UI EVENTS: Initializing event handlers for resize & cell interaction");
            
            // Column resize state management
            _isResizing = false;
            _resizingColumnIndex = -1;
            _resizeStartX = 0;
            _resizeStartWidth = 0;
            
            // Event handlers will be attached to ItemsRepeater elements dynamically
            // when UI elements are rendered (in DataGridUIManager)
            
            _logger?.Info("✅ UI EVENTS: Event handlers initialized successfully");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI EVENTS ERROR: Failed to initialize event handlers");
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

            _logger?.Info("🎨 UI SIZING: Applied sizing constraints");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: ApplyUISizing failed");
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
            _logger?.Info("🎨 UI COLOR: Applied color configuration to UI elements");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: ApplyColorConfiguration failed");
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
            _logger?.Info("🎨 UI VIRTUAL: Virtualization initialized");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: InitializeUIVirtualizationAsync failed");
        }
    }

    /// <summary>
    /// Render all cells using proper UI Manager - kvalitné riešenie s comprehensive error logging
    /// </summary>
    private async Task RenderAllCellsAsync()
    {
        if (_uiManager == null)
        {
            _logger?.Error("🚨 UI ERROR: UIManager is null, cannot render cells");
            throw new InvalidOperationException("UIManager must be initialized before rendering cells");
        }

        try
        {
            _logger?.Info("🎨 UI RENDER: Starting comprehensive cell rendering via UIManager...");
            
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
                        _logger?.Info("✅ UI RENDER: Comprehensive rendering completed on UI thread - Headers: {HeaderCount}, Rows: {RowCount}, Cells: {CellCount}", 
                            stats.HeaderCount, stats.RowCount, stats.TotalCellCount);
                        
                        completionSource.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        _logger?.Error(ex, "🚨 UI RENDER ERROR: RefreshAllUIAsync failed on UI thread");
                        completionSource.SetException(ex);
                    }
                });
                
                // Wait for UI thread operation to complete
                await completionSource.Task;
            }
            else
            {
                // Fallback to direct call if no dispatcher available
                _logger?.Warning("⚠️ UI RENDER: No DispatcherQueue available, using direct call");
                await _uiManager.RefreshAllUIAsync();
                
                var stats = _uiManager.GetRenderingStats();
                _logger?.Info("✅ UI RENDER: Comprehensive rendering completed (direct) - Headers: {HeaderCount}, Rows: {RowCount}, Cells: {CellCount}", 
                    stats.HeaderCount, stats.RowCount, stats.TotalCellCount);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: Comprehensive cell rendering failed");
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
            _logger?.Error("🚨 UI ERROR: UIManager is null, cannot update validation visuals");
            return;
        }

        try
        {
            await _uiManager.UpdateValidationUIAsync();
            _logger?.Info("✅ UI VALIDATION: Validation visuals updated via UIManager");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateValidationVisualsAsync failed");
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
            _logger?.Error("🚨 UI ERROR: UIManager is null, cannot update row UI");
            return;
        }

        try
        {
            await _uiManager.UpdateRowUIAsync(rowIndex);
            _logger?.Info("✅ UI ROW: Updated row {Row} UI via UIManager", rowIndex);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateSpecificRowUIAsync failed for row {Row}", rowIndex);
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
            _logger?.Info("🎨 UI CELL: Updated cell [{Row},{Column}] UI", row, column);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateSpecificCellUIAsync failed for cell [{Row},{Column}]", row, column);
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
            _logger?.Info("🎨 UI COLUMN: Updated column {Column} UI", columnName);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 UI ERROR: UpdateSpecificColumnUIAsync failed for column {Column}", columnName);
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
    // Simplified drag tracking - just start and end cells
    private DataCellModel? _dragStartCell = null;
    private DataCellModel? _dragEndCell = null;

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
                
                // Make handle slightly visible for debugging
                handle.Fill = new SolidColorBrush(Microsoft.UI.Colors.LightGray) { Opacity = 0.3 };
                
                _logger?.Info("🔍 RESIZE: Pointer entered ±2px resize area - cursor changed");
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to handle pointer enter");
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
                // Reset cursor to default and handle appearance
                this.ProtectedCursor = null;
                handle.Fill = new SolidColorBrush(Microsoft.UI.Colors.Transparent);
                _logger?.Info("🔍 RESIZE: Pointer exited resize handle");
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to handle pointer exit");
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
                    
                    _logger?.Debug("🔄 RESIZE: Started resizing column {ColumnIndex} from width {StartWidth}", 
                        columnIndex, _resizeStartWidth);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to start resize operation");
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
                    
                    _logger?.Debug("🔄 RESIZE: Column {ColumnIndex} width = {NewWidth} (delta: {Delta})", 
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
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to handle pointer move");
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
                
                _logger?.Info("✅ RESIZE: Completed resizing column {ColumnIndex} to width {FinalWidth}", 
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
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to complete resize operation");
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
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to get column index from resize handle");
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
            
            _logger?.Debug("⚡ RESIZE FAST: Updated header {ColumnIndex} width to {NewWidth}", 
                columnIndex, newWidth);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to fast update column width");
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
            
            _logger?.Debug("🔄 RESIZE: Updated column {ColumnIndex} width to {NewWidth} across all rows", 
                columnIndex, newWidth);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to update column width");
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

            _logger?.Debug("🔒 RESIZE CONSTRAINTS: Column {ColumnIndex} - Target: {Target}, Min: {Min}, Max: {Max}, Final: {Final}",
                columnIndex, targetWidth, minWidth, maxWidth == double.MaxValue ? "∞" : maxWidth.ToString(), constrainedWidth);

            return constrainedWidth;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RESIZE ERROR: Failed to apply width constraints for column {ColumnIndex}", columnIndex);
            // Fallback: Use reasonable default constraints
            return Math.Max(50, Math.Min(targetWidth, 1000));
        }
    }

    #endregion

    #region Cell Editing Event Handlers

    // REMOVED: DisplayArea_Tapped - jeden klik má robiť iba focus, nie edit mode
    // Edit mode sa spúšťa iba pri dvojkliku alebo type-to-edit

    /// <summary>
    /// Event handler for real-time validation during text changes
    /// </summary>
    private async void EditTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        try
        {
            if (sender is TextBox textBox && textBox.DataContext is DataCellModel cellModel)
            {
                // Update the DisplayText immediately for real-time validation and UI refresh
                cellModel.DisplayText = textBox.Text;
                
                // REAL-TIME VALIDATION: Trigger validation immediately on UI thread
                DispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, async () =>
                {
                    try
                    {
                        // Perform real-time validation and update visual state
                        await PerformRealTimeValidationAsync(cellModel);
                        
                        // AUTO-RESIZE: Recalculate row height when content changes
                        await UpdateRowHeightAsync(cellModel.RowIndex);
                    }
                    catch (Exception ex)
                    {
                        _logger?.Error(ex, "🚨 REAL-TIME ERROR: Failed real-time validation for cell [{Row},{Col}]", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                });
                
                _logger?.Debug("📝 TEXT CHANGED: Real-time validation and UI refresh triggered for cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 TEXT CHANGE ERROR: Failed to handle text change");
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
                _logger?.Info("💾 CELL EDIT: Exited edit mode for cell [{Row},{Col}] with value '{Value}'", 
                    cellModel.RowIndex, cellModel.ColumnIndex, cellModel.DisplayText);

                // Save the value to the controller
                await SaveCellValueAsync(cellModel);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CELL EDIT ERROR: Failed to exit edit mode");
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
                            _logger?.Debug("📝 SHIFT+ENTER: Inserted newline at position {Position} in cell [{Row},{Col}]", 
                                cursorPosition, cellModel.RowIndex, cellModel.ColumnIndex);
                        }
                        else
                        {
                            // Enter: Save and exit edit mode
                            cellModel.IsEditing = false;
                            await SaveCellValueAsync(cellModel);
                            e.Handled = true;
                            _logger?.Info("✅ CELL EDIT: Enter pressed - saved cell [{Row},{Col}]", 
                                cellModel.RowIndex, cellModel.ColumnIndex);
                        }
                        break;
                        
                    case Windows.System.VirtualKey.Escape:
                        // Cancel edit without saving
                        cellModel.IsEditing = false;
                        // Restore original value
                        cellModel.DisplayText = cellModel.Value?.ToString() ?? string.Empty;
                        e.Handled = true;
                        _logger?.Info("❌ CELL EDIT: Escape pressed - cancelled edit for cell [{Row},{Col}]", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                        break;
                        
                    case Windows.System.VirtualKey.Tab:
                        // Tab in edit mode: Insert tab character at cursor position
                        int tabCursorPosition = textBox.SelectionStart;
                        string tabCurrentText = textBox.Text ?? string.Empty;
                        string tabNewText = tabCurrentText.Insert(tabCursorPosition, "\t");
                        
                        textBox.Text = tabNewText;
                        textBox.SelectionStart = tabCursorPosition + 1; // Move cursor after the tab
                        
                        e.Handled = true;
                        _logger?.Debug("📝 TAB INSERT: Inserted tab at position {Position} in cell [{Row},{Col}]", 
                            tabCursorPosition, cellModel.RowIndex, cellModel.ColumnIndex);
                        break;
                        
                    case Windows.System.VirtualKey.C:
                        // FIXED: Handle Ctrl+C in edit mode by forwarding to main handler
                        bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control)
                            .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                        if (isCtrlPressed)
                        {
                            // Commit current edit and then handle copy
                            await CommitCellEditAsync(stayOnCell: true);
                            await CopySelectedCellsAsync();
                            e.Handled = true;
                            _logger?.Info("📋 COPY FROM EDIT: Committed edit and copied selected cells");
                        }
                        break;
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CELL EDIT ERROR: Failed to handle key press");
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
            _logger?.Error(ex, "🚨 TEXTBOX SEARCH ERROR: Failed to find TextBox for cell [{Row},{Col}]", 
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
            _logger?.Error(ex, "🚨 CELL TEXTBOX SEARCH ERROR: Failed to find TextBox in row for column {Column}", columnIndex);
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
                
                _logger?.Info("📝 TYPE-TO-EDIT: Cell [{Row},{Col}] entered edit mode by typing '{Key}' (TextBox will handle character)", 
                    focusedCell.RowIndex, focusedCell.ColumnIndex, key);
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 TYPE-TO-EDIT ERROR: Failed to check for type-to-edit");
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
    /// Trigger UI refresh for internal method calls (not public API)
    /// </summary>
    private async Task TriggerInternalUIRefreshAsync(string reason)
    {
        try
        {
            _logger?.Debug("🔄 INTERNAL REFRESH: {Reason}", reason);
            
            // FOCUS RETENTION: Save current editing state before refresh
            var wasEditing = _currentEditingCell != null && _currentEditingCell.IsEditing;
            var editingRowIndex = _currentEditingCell?.RowIndex ?? -1;
            var editingColumnIndex = _currentEditingCell?.ColumnIndex ?? -1;
            
            // Use DispatcherQueue to ensure UI thread execution
            if (DispatcherQueue != null)
            {
                DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, async () =>
                {
                    try
                    {
                        await RenderAllCellsAsync();
                        
                        // FOCUS RETENTION: Restore focus and editing state after refresh
                        if (wasEditing && editingRowIndex >= 0 && editingColumnIndex >= 0)
                        {
                            await RestoreFocusAfterRefreshAsync(editingRowIndex, editingColumnIndex);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.Error(ex, "🚨 INTERNAL REFRESH ERROR: Failed to refresh UI - {Reason}", reason);
                    }
                });
            }
            else
            {
                // Fallback: direct call if DispatcherQueue is not available
                await RenderAllCellsAsync();
                
                // FOCUS RETENTION: Restore focus and editing state after refresh
                if (wasEditing && editingRowIndex >= 0 && editingColumnIndex >= 0)
                {
                    await RestoreFocusAfterRefreshAsync(editingRowIndex, editingColumnIndex);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 INTERNAL REFRESH ERROR: Failed to trigger UI refresh - {Reason}", reason);
        }
    }

    /// <summary>
    /// Restore focus and editing state after UI refresh
    /// </summary>
    private async Task RestoreFocusAfterRefreshAsync(int rowIndex, int columnIndex)
    {
        try
        {
            // Find the cell that was being edited
            var cellModel = FindCellModel(rowIndex, columnIndex);
            if (cellModel != null)
            {
                // Restore focus
                await MoveFocusToAsync(rowIndex, columnIndex);
                
                // Restore editing state if needed
                if (!cellModel.IsEditing)
                {
                    StartCellEditing(cellModel, false); // Keep current content
                }
                
                _logger?.Info("🎯 FOCUS RESTORED: Cell [{Row},{Col}] focus and editing state restored after refresh", 
                    rowIndex, columnIndex);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 FOCUS RESTORE ERROR: Failed to restore focus after refresh");
        }
    }

    /// <summary>
    /// Perform real-time validation and update visual state immediately
    /// </summary>
    private async Task PerformRealTimeValidationAsync(DataCellModel cellModel)
    {
        try
        {
            // First save the current value to trigger validation
            if (_controller != null)
            {
                object? newValue = ConvertDisplayTextToValue(cellModel.DisplayText, cellModel.ColumnName);
                await _controller.SetCellValueAsync(cellModel.RowIndex, cellModel.ColumnIndex, newValue);
                cellModel.Value = newValue;
            }
            
            // Perform validation update
            if (_uiManager != null)
            {
                await _uiManager.UpdateCellValidationAsync(cellModel);
            }
            
            // Update ValidationAlerts column if it exists
            await UpdateValidationAlertsColumnAsync(cellModel.RowIndex);
            
            _logger?.Debug("🔍 REAL-TIME VALIDATION: Updated validation for cell [{Row},{Col}] with value '{Value}'", 
                cellModel.RowIndex, cellModel.ColumnIndex, cellModel.DisplayText);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 REAL-TIME VALIDATION ERROR: Failed to validate cell [{Row},{Col}]", 
                cellModel.RowIndex, cellModel.ColumnIndex);
        }
    }

    /// <summary>
    /// Update ValidationAlerts column for specific row
    /// </summary>
    private async Task UpdateValidationAlertsColumnAsync(int rowIndex)
    {
        try
        {
            if (_uiManager?.RowsCollection != null && rowIndex >= 0 && rowIndex < _uiManager.RowsCollection.Count)
            {
                var row = _uiManager.RowsCollection[rowIndex];
                var validationCell = row.Cells.FirstOrDefault(c => c.ColumnName == "ValidationAlerts");
                
                if (validationCell != null && _controller != null)
                {
                    // Get validation errors for this row
                    var errors = new List<string>();
                    foreach (var cell in row.Cells.Where(c => c.ColumnName != "ValidationAlerts"))
                    {
                        if (!string.IsNullOrEmpty(cell.ValidationError))
                        {
                            errors.Add($"{cell.ColumnName}: {cell.ValidationError}");
                        }
                    }
                    
                    // Update ValidationAlerts cell
                    string alertText = errors.Count > 0 ? string.Join("; ", errors) : "";
                    validationCell.DisplayText = alertText;
                    validationCell.Value = alertText;
                    
                    // Save to controller
                    await _controller.SetCellValueAsync(rowIndex, validationCell.ColumnIndex, alertText);
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 VALIDATION ALERTS ERROR: Failed to update ValidationAlerts column for row {Row}", rowIndex);
        }
    }

    /// <summary>
    /// Find cell model by row and column index
    /// </summary>
    private DataCellModel? FindCellModel(int rowIndex, int columnIndex)
    {
        try
        {
            if (_uiManager?.RowsCollection != null && 
                rowIndex >= 0 && rowIndex < _uiManager.RowsCollection.Count)
            {
                var row = _uiManager.RowsCollection[rowIndex];
                if (columnIndex >= 0 && columnIndex < row.Cells.Count)
                {
                    return row.Cells[columnIndex];
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 FIND CELL ERROR: Failed to find cell [{Row},{Col}]", rowIndex, columnIndex);
            return null;
        }
    }

    /// <summary>
    /// Update row height when cell content changes (auto-resize functionality)
    /// </summary>
    private async Task UpdateRowHeightAsync(int rowIndex)
    {
        try
        {
            if (_uiManager?.RowsCollection != null && 
                rowIndex >= 0 && rowIndex < _uiManager.RowsCollection.Count)
            {
                var rowModel = _uiManager.RowsCollection[rowIndex];
                
                // Calculate new height based on current cell content
                double maxHeight = 32; // Minimum row height
                
                foreach (var cell in rowModel.Cells)
                {
                    // Use UIManager's calculation method through reflection or add public method
                    var cellHeight = await CalculateCellHeight(cell.DisplayText, cell.Width);
                    maxHeight = Math.Max(maxHeight, cellHeight);
                }
                
                // Update row and all cells in the row
                rowModel.Height = maxHeight;
                foreach (var cell in rowModel.Cells)
                {
                    cell.Height = maxHeight;
                }
                
                _logger?.Debug("📏 ROW HEIGHT UPDATED: Row {Row} height updated to {Height}", 
                    rowIndex, maxHeight);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW HEIGHT UPDATE ERROR: Failed to update height for row {Row}", rowIndex);
        }
    }

    /// <summary>
    /// Calculate required height for cell content
    /// </summary>
    private async Task<double> CalculateCellHeight(string text, double maxWidth, double minWidth = 60)
    {
        try
        {
            // Ensure minimum column width is respected
            var availableWidth = Math.Max(maxWidth, minWidth);
            
            // Account for cell padding (6px left + 6px right = 12px total)
            var textWidth = availableWidth - 12;
            
            if (string.IsNullOrEmpty(text) || textWidth <= 0)
            {
                return 32; // Minimum row height
            }

            // Use DispatcherQueue to ensure UI thread execution for measurements
            var completionSource = new TaskCompletionSource<double>();
            
            DispatcherQueue?.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
            {
                try
                {
                    // Create a TextBlock to measure text
                    var textBlock = new Microsoft.UI.Xaml.Controls.TextBlock
                    {
                        Text = text,
                        TextWrapping = Microsoft.UI.Xaml.TextWrapping.Wrap,
                        Width = textWidth,
                        FontSize = 14, // Default font size
                        FontFamily = new Microsoft.UI.Xaml.Media.FontFamily("Segoe UI"), // Default font
                        Padding = new Microsoft.UI.Xaml.Thickness(0),
                        Margin = new Microsoft.UI.Xaml.Thickness(0)
                    };

                    // Measure the text
                    textBlock.Measure(new Windows.Foundation.Size(textWidth, double.PositiveInfinity));
                    var measuredHeight = textBlock.DesiredSize.Height;
                    
                    // Add padding (top + bottom = 4px total) and ensure minimum height
                    var requiredHeight = Math.Max(measuredHeight + 8, 32);
                    
                    completionSource.SetResult(requiredHeight);
                }
                catch (Exception ex)
                {
                    completionSource.SetException(ex);
                }
            });
            
            return await completionSource.Task;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CELL HEIGHT CALC ERROR: Failed to calculate height for text");
            return 32; // Fallback to minimum height
        }
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
            
            _logger?.Info("📝 EDIT START: Cell [{Row},{Col}] editing started (clear: {Clear}, original: '{Original}')", 
                cellModel.RowIndex, cellModel.ColumnIndex, clearContent, _originalEditValue);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 EDIT START ERROR: Failed to start cell editing");
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
                _logger?.Error("🚨 SAVE ERROR: CellModel is null");
                return;
            }
            
            if (_controller == null)
            {
                _logger?.Error("🚨 SAVE ERROR: Controller is null");
                return;
            }
            
            // Convert display text to appropriate type
            object? newValue = ConvertDisplayTextToValue(cellModel.DisplayText, cellModel.ColumnName);
            
            _logger?.Error("🚨💾 DEBUG SAVE START: Saving cell [{Row},{Col}] '{OldValue}' → '{NewValue}' (DisplayText: '{DisplayText}')", 
                cellModel.RowIndex, cellModel.ColumnIndex, cellModel.Value?.ToString() ?? "null", newValue?.ToString() ?? "null", cellModel.DisplayText);
                
            // CRITICAL: Save to controller first (this persists to underlying TableCore)
            await _controller.SetCellValueAsync(cellModel.RowIndex, cellModel.ColumnIndex, newValue);
            
            _logger?.Error("🚨💾 DEBUG SAVE CORE: Successfully saved to TableCore [{Row},{Col}] = '{Value}'", 
                cellModel.RowIndex, cellModel.ColumnIndex, newValue?.ToString() ?? "null");
            
            // CRITICAL: Update the model with the actual saved value to maintain consistency
            cellModel.Value = newValue;
            
            // PERSISTENCE FIX: Verify the data was actually saved by reading it back from TableCore
            try
            {
                // Wait a bit to ensure save is complete
                await Task.Delay(10);
                
                var verifyValue = await _controller.TableCore.GetCellValueAsync(cellModel.RowIndex, cellModel.ColumnIndex);
                _logger?.Error("🚨💾 DEBUG VERIFY: Read back value [{Row},{Col}] = '{VerifyValue}' (expected: '{ExpectedValue}')", 
                    cellModel.RowIndex, cellModel.ColumnIndex, verifyValue?.ToString() ?? "null", newValue?.ToString() ?? "null");
                
                // Check if verification matches
                bool valuesMatch = (verifyValue?.ToString() ?? "") == (newValue?.ToString() ?? "");
                _logger?.Error("🚨💾 DEBUG VERIFY RESULT: Values match = {Match}", valuesMatch);
                
                // CRITICAL FIX: Only update Value, preserve DisplayText to avoid UI reversion
                cellModel.Value = verifyValue;
                
                // Only update DisplayText if the values don't match (indicating a conversion/save issue)
                if (!valuesMatch)
                {
                    _logger?.Error("🚨💾 MISMATCH: Updating DisplayText due to save/verify mismatch");
                    cellModel.DisplayText = verifyValue?.ToString() ?? string.Empty;
                }
                else
                {
                    _logger?.Error("✅💾 MATCH: Keeping original DisplayText, values match");
                    // Keep the original DisplayText that user entered
                }
            }
            catch (Exception verifyEx)
            {
                _logger?.Error(verifyEx, "🚨 CELL VERIFY ERROR: Failed to verify saved value [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
            
            // Update validation for this specific cell
            if (_uiManager != null)
            {
                await _uiManager.UpdateCellValidationAsync(cellModel);
            }
            
            _logger?.Error("🚨💾 DEBUG SAVE COMPLETE: Cell [{Row},{Col}] save process completed with verification", 
                cellModel.RowIndex, cellModel.ColumnIndex);
            
            // Clear editing state after successful save (commit)
            if (_currentEditingCell == cellModel)
            {
                _currentEditingCell = null;
                _originalEditValue = null;
                _logger?.Debug("📝 EDIT COMMIT: Cell editing state cleared after save");
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CELL SAVE ERROR: Failed to save cell [{Row},{Col}]", 
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
            _logger?.Error(ex, "⚠️ CELL CONVERT: Failed to convert '{DisplayText}' for column '{Column}', using string fallback", 
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
                
                _logger?.Debug("🎯 CELL SELECT: Cell [{Row},{Col}] selected (Ctrl: {Ctrl}, Shift: {Shift})", 
                    cellModel.RowIndex, cellModel.ColumnIndex, isCtrlPressed, isShiftPressed);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CELL SELECT ERROR: Failed to select cell");
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
                _logger?.Debug("🖱️ CELL CONTEXT: Right-clicked cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CELL CONTEXT ERROR: Failed to handle right-click");
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
            _logger?.Error(ex, "🚨 SELECTION ERROR: Failed to clear cell selection");
        }
    }

    /// <summary>
    /// Update visual styling for cell focus/selection/copy state using proper color configuration
    /// </summary>
    private void UpdateCellSelectionVisuals(DataCellModel cellModel)
    {
        try
        {
            // Determine the appropriate background and border colors based on cell state priority:
            // 1. Copy mode (highest priority)
            // 2. Focus/Selection
            // 3. Validation error
            // 4. Normal state
            
            if (cellModel.IsCopied)
            {
                // Copy mode - use configured copy mode colors
                var copyBackgroundColor = _controller?.ColorConfig?.CopyModeBackgroundColor ?? 
                    Windows.UI.Color.FromArgb(100, 173, 216, 230); // Default light blue
                cellModel.BackgroundBrush = CreateBrush(copyBackgroundColor);
                cellModel.BorderBrush = CreateBrush(_controller?.ColorConfig?.CellBorderColor ?? Microsoft.UI.Colors.LightGray);
                
                _logger?.Debug("🎨 COPY STYLE: Applied copy mode background to cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
            else if (cellModel.IsSelected || cellModel.IsFocused)
            {
                // Focus/Selection state - use configured selection colors
                var selectionBackgroundColor = _controller?.ColorConfig?.SelectionBackgroundColor ?? 
                    Windows.UI.Color.FromArgb(100, 144, 238, 144); // Default light green
                var focusRingColor = _controller?.ColorConfig?.FocusRingColor ?? 
                    Windows.UI.Color.FromArgb(255, 144, 238, 144); // Default green border
                    
                cellModel.BackgroundBrush = CreateBrush(selectionBackgroundColor);
                cellModel.BorderBrush = CreateBrush(focusRingColor);
                
                _logger?.Debug("🎨 FOCUS STYLE: Applied focus/selection background to cell [{Row},{Col}]", 
                    cellModel.RowIndex, cellModel.ColumnIndex);
            }
            else if (!cellModel.IsValid)
            {
                // Validation error state
                cellModel.BorderBrush = CreateBrush(_controller?.ColorConfig?.ValidationErrorBorderColor ?? Microsoft.UI.Colors.Red);
                cellModel.BackgroundBrush = CreateBrush(_controller?.ColorConfig?.ValidationErrorBackgroundColor ?? Microsoft.UI.Colors.LightPink);
            }
            else
            {
                // Normal state - restore original colors
                cellModel.BorderBrush = CreateBrush(_controller?.ColorConfig?.CellBorderColor ?? Microsoft.UI.Colors.LightGray);
                cellModel.BackgroundBrush = CreateBrush(_controller?.ColorConfig?.CellBackgroundColor ?? Microsoft.UI.Colors.White);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 VISUAL UPDATE ERROR: Failed to update cell selection visuals");
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
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // COMPREHENSIVE LOGGING: Method entry with sender information
            _logger?.Info("🎯 POINTER PRESSED: CellBorder_PointerPressed started - Sender: {SenderType}", 
                sender?.GetType()?.Name ?? "null");
            _logger?.Debug("🎯 POINTER PRESSED DEBUG: Thread: {ThreadId}, HasPointer: {HasPointer}", 
                Environment.CurrentManagedThreadId, e?.Pointer != null);
            
            if (sender is Border border && border.DataContext is DataCellModel cellModel)
            {
                // COMPLETE USER ACTION LOGGING
                _logger?.Info("👆 USER CLICK: Cell [{Row},{Col}] = '{Text}' clicked", 
                    cellModel.RowIndex, cellModel.ColumnIndex, 
                    cellModel.DisplayText?.Length > 100 ? cellModel.DisplayText?.Substring(0, 100) + "..." : cellModel.DisplayText);
                _logger?.Debug("👆 USER CLICK DEBUG: Cell state - IsSelected: {IsSelected}, IsFocused: {IsFocused}, IsEditing: {IsEditing}", 
                    cellModel.IsSelected, cellModel.IsFocused, cellModel.IsEditing);
                    
                // Check for Shift key for range selection
                bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift)
                    .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                    
                bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control)
                    .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
                    
                // LOG COMPLETE INTERACTION CONTEXT
                _logger?.Info("⌨️ MODIFIERS: Ctrl={Ctrl}, Shift={Shift}", isCtrlPressed, isShiftPressed);
                
                // LOG CURRENT SELECTION STATE BEFORE ACTION
                var currentSelectedCells = GetAllSelectedCells();
                _logger?.Info("🔍 SELECTION BEFORE: {SelectedCount} cells selected", currentSelectedCells.Count);
                _logger?.Debug("🔍 SELECTION BEFORE DEBUG: Selected cells: {SelectedCells}", 
                    string.Join(", ", currentSelectedCells.Select(c => $"[{c.RowIndex},{c.ColumnIndex}]")));

                // Check if it's left mouse button and no resize operation
                if (e.Pointer.PointerDeviceType == Microsoft.UI.Input.PointerDeviceType.Mouse && 
                    e.GetCurrentPoint(border).Properties.IsLeftButtonPressed && !_isResizing)
                {
                    // SIMPLE CLICK LOGIC: 
                    // 1. End any active edit mode first
                    if (_currentEditingCell != null && _currentEditingCell.IsEditing)
                    {
                        await CommitCellEditAsync(stayOnCell: false);
                    }
                    
                    // 2. Clear all previous selections for normal click
                    if (!isCtrlPressed)
                    {
                        _logger?.Info("🧹 SELECTION CLEAR: Normal click - clearing all previous selections");
                        ClearAllSelection();
                        _logger?.Info("✅ SELECTION CLEAR: All previous selections cleared");
                    }
                    else
                    {
                        _logger?.Info("🎯 CTRL+CLICK: Preserving existing selection, adding to multi-select");
                    }
                    
                    // 3. Select this cell and set focus
                    cellModel.IsSelected = true;
                    UpdateCellSelectionVisuals(cellModel); // CRITICAL: Update visual styling
                    _logger?.Info("✅ CELL SELECTED: Cell [{Row},{Col}] marked as selected", cellModel.RowIndex, cellModel.ColumnIndex);
                    
                    // CTRL+CLICK FIX: For Ctrl+Click, update focus without calling MoveFocusToAsync to preserve selection
                    if (isCtrlPressed)
                    {
                        // Update focus tracking variables directly without clearing other cells
                        _focusedRowIndex = cellModel.RowIndex;
                        _focusedColumnIndex = cellModel.ColumnIndex;
                        if (_focusedCell != null)
                        {
                            _focusedCell.IsFocused = false;
                        }
                        _focusedCell = cellModel;
                        cellModel.IsFocused = true;
                        _logger?.Info("🎯 CTRL+CLICK FOCUS: Focus updated directly to preserve selection - Cell [{Row},{Col}]", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                    else
                    {
                        // Normal click - use MoveFocusToAsync
                        await MoveFocusToAsync(cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                    
                    // ENHANCED FOCUS MANAGEMENT: Aggressive focus acquisition for keyboard events
                    await EnsureKeyboardFocusAsync();
                    
                    _logger?.Info("🎯 FOCUS STATE: Enhanced focus completed - Current FocusState = {CurrentFocusState}", 
                        this.FocusState);
                    
                    // 4. Prepare for drag selection (but NOT for Ctrl+Click)
                    if (!isCtrlPressed)
                    {
                        _dragStartCell = cellModel;
                        _logger?.Info("🎯 DRAG PREPARED: Drag selection prepared for normal click");
                    }
                    else
                    {
                        _dragStartCell = null; // Disable drag for Ctrl+Click to preserve multi-select
                        _logger?.Info("🎯 CTRL+CLICK: Drag selection disabled to preserve multi-select");
                    }
                    
                    // 5. IMPROVED EDIT MODE LOGIC: Second click on focused cell starts editing
                    bool willStartEditing = false;
                    
                    // Check if this cell is already focused (second click)
                    if (_focusedCell == cellModel && cellModel.IsFocused && !cellModel.IsReadOnly)
                    {
                        // Second click on already focused cell = start editing
                        willStartEditing = true;
                        StartCellEditing(cellModel, false);
                        _logger?.Info("📝 EDIT MODE START: Second click on focused cell [{Row},{Col}] - entering edit mode", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                    else
                    {
                        // First click or click on different cell = just focus
                        _logger?.Info("🎯 FOCUS ONLY: First click or different cell [{Row},{Col}] - focus only", 
                            cellModel.RowIndex, cellModel.ColumnIndex);
                    }
                    
                    _logger?.Info("🎯 CLICK RESULT: Cell [{Row},{Col}] selected, EditMode: {EditMode}", 
                        cellModel.RowIndex, cellModel.ColumnIndex, willStartEditing);
                    
                    // LOG FINAL SELECTION STATE AFTER ACTION
                    var finalSelectedCells = GetAllSelectedCells();
                    _logger?.Info("🔍 SELECTION AFTER: {SelectedCount} cells selected after click action", finalSelectedCells.Count);
                    _logger?.Debug("🔍 SELECTION AFTER DEBUG: Selected cells: {SelectedCells}", 
                        string.Join(", ", finalSelectedCells.Select(c => $"[{c.RowIndex},{c.ColumnIndex}]")));
                    
                    // Capture pointer on DataGrid for global tracking
                    this.CapturePointer(e.Pointer);
                }
            }
            
            // COMPREHENSIVE LOGGING: Method exit with timing
            stopwatch.Stop();
            _logger?.Info("✅ POINTER PRESSED: CellBorder_PointerPressed completed successfully in {ElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds);
            _logger?.Debug("✅ POINTER PRESSED DEBUG: Final focus state - Control: {ControlFocus}, FocusedCell: [{Row},{Col}]", 
                this.FocusState, _focusedRowIndex, _focusedColumnIndex);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.Error(ex, "🚨 CLICK ERROR: CellBorder_PointerPressed failed after {ElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds);
        }
    }

    /// <summary>
    /// Global pointer moved handler for drag selection across multiple cells
    /// </summary>
    private void DataGrid_PointerMoved(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            // SIMPLE DRAG: If we have a drag start cell, update selection
            if (_dragStartCell != null)
            {
                var currentCell = FindCellUnderPointer(e.GetCurrentPoint(this).Position);
                if (currentCell != null)
                {
                    // Only update if we moved to a different cell
                    if (_dragEndCell == null || 
                        _dragEndCell.RowIndex != currentCell.RowIndex || 
                        _dragEndCell.ColumnIndex != currentCell.ColumnIndex)
                    {
                        _dragEndCell = currentCell;
                        
                        // IMMEDIATE update without UI thread delay for better responsiveness
                        UpdateSimpleDragSelection();
                        
                        _logger?.Debug("🎯 DRAG MOVED: From [{StartRow},{StartCol}] to [{EndRow},{EndCol}]", 
                            _dragStartCell.RowIndex, _dragStartCell.ColumnIndex, _dragEndCell.RowIndex, _dragEndCell.ColumnIndex);
                    }
                }
                else
                {
                    _logger?.Debug("🎯 DRAG: No cell found under pointer");
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 DRAG ERROR: Failed to update drag selection");
        }
    }

    /// <summary>
    /// Global pointer released handler for drag selection completion
    /// </summary>
    private async void DataGrid_PointerReleased(object sender, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs e)
    {
        try
        {
            // SIMPLE LOGIC: Just release pointer and reset variables
            this.ReleasePointerCapture(e.Pointer);
            
            // If we had a drag operation, log completion
            if (_dragStartCell != null && _dragEndCell != null)
            {
                _logger?.Info("✅ DRAG COMPLETE: Selected from [{StartRow},{StartCol}] to [{EndRow},{EndCol}]", 
                    _dragStartCell.RowIndex, _dragStartCell.ColumnIndex, _dragEndCell.RowIndex, _dragEndCell.ColumnIndex);
                
                // CRITICAL FIX: Don't call MoveFocusToAsync as it might interfere with selection
                // Just update focus tracking variables directly
                _focusedRowIndex = _dragEndCell.RowIndex;
                _focusedColumnIndex = _dragEndCell.ColumnIndex;
                _focusedCell = _dragEndCell;
                
                _logger?.Info("🎯 DRAG FOCUS: Focus set to end cell [{Row},{Col}] without clearing selection", 
                    _dragEndCell.RowIndex, _dragEndCell.ColumnIndex);
            }
            
            // Reset drag variables
            _dragStartCell = null;
            _dragEndCell = null;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 DRAG ERROR: Failed to complete drag selection");
        }
    }


    /// <summary>
    /// Find cell under the given pointer position
    /// </summary>
    private DataCellModel? FindCellUnderPointer(Windows.Foundation.Point position)
    {
        try
        {
            if (_uiManager == null) return null;

            _logger?.Debug("🎯 HIT TEST: Finding cell at position ({X}, {Y})", position.X, position.Y);

            // IMPROVED: Use mathematical approach based on cell dimensions
            // Get header height (typically 32px)
            double headerHeight = 32.0;
            
            // Account for header - subtract header height from Y position
            double adjustedY = position.Y - headerHeight;
            
            if (adjustedY < 0)
            {
                _logger?.Debug("🎯 HIT TEST: Position is in header area, no cell found");
                return null; // Click is in header area
            }

            // Calculate row index based on Y position and row height
            double rowHeight = 32.0; // Default row height
            int rowIndex = (int)(adjustedY / rowHeight);
            
            // Calculate column index based on X position and column widths
            int columnIndex = -1;
            double currentX = 0;
            
            for (int col = 0; col < ColumnCount; col++)
            {
                double columnWidth = 100.0; // Default column width - should get from actual columns
                if (_uiManager.HeadersCollection.Count > col)
                {
                    columnWidth = _uiManager.HeadersCollection[col].Width;
                }
                
                if (position.X >= currentX && position.X < currentX + columnWidth)
                {
                    columnIndex = col;
                    break;
                }
                currentX += columnWidth;
            }

            _logger?.Debug("🎯 HIT TEST: Calculated position - Row: {Row}, Column: {Col} (adjustedY: {AdjY}, rowHeight: {RH})", 
                rowIndex, columnIndex, adjustedY, rowHeight);

            // Validate bounds and get cell
            if (rowIndex >= 0 && rowIndex < _uiManager.RowsCollection.Count && 
                columnIndex >= 0 && columnIndex < ColumnCount)
            {
                var row = _uiManager.RowsCollection[rowIndex];
                if (columnIndex < row.Cells.Count)
                {
                    var cell = row.Cells[columnIndex];
                    _logger?.Debug("🎯 HIT TEST: Found cell [{Row},{Col}] = '{Value}'", 
                        cell.RowIndex, cell.ColumnIndex, cell.DisplayText);
                    return cell;
                }
            }

            _logger?.Debug("🎯 HIT TEST: No valid cell found for position ({X}, {Y})", position.X, position.Y);
            return null;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 HIT TEST ERROR: Failed to find cell under pointer at ({X}, {Y})", position.X, position.Y);
            return null;
        }
    }

    /// <summary>
    /// Simple drag selection update - just select cells in rectangle
    /// </summary>
    private void UpdateSimpleDragSelection()
    {
        try
        {
            if (_dragStartCell == null || _dragEndCell == null || _uiManager == null) return;

            // Calculate selection bounds
            int startRow = Math.Min(_dragStartCell.RowIndex, _dragEndCell.RowIndex);
            int endRow = Math.Max(_dragStartCell.RowIndex, _dragEndCell.RowIndex);
            int startCol = Math.Min(_dragStartCell.ColumnIndex, _dragEndCell.ColumnIndex);
            int endCol = Math.Max(_dragStartCell.ColumnIndex, _dragEndCell.ColumnIndex);

            // CRITICAL FIX: Run selection update on UI thread for immediate visual feedback
            if (DispatcherQueue != null)
            {
                DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, () =>
                {
                    try
                    {
                        // Clear all selections first
                        ClearAllCellSelection();

                        // Select cells in rectangle and give them background
                        foreach (var row in _uiManager.RowsCollection)
                        {
                            if (row.RowIndex >= startRow && row.RowIndex <= endRow)
                            {
                                for (int colIndex = startCol; colIndex <= endCol && colIndex < row.Cells.Count; colIndex++)
                                {
                                    var cell = row.Cells[colIndex];
                                    cell.IsSelected = true;
                                    UpdateCellSelectionVisuals(cell); // Give it focus background
                                    
                                    _logger?.Debug("🎯 DRAG SELECT: Cell [{Row},{Col}] selected", cell.RowIndex, cell.ColumnIndex);
                                }
                            }
                        }

                        _logger?.Info("🎯 SIMPLE DRAG: Selected rectangle from [{StartRow},{StartCol}] to [{EndRow},{EndCol}] on UI thread", 
                            startRow, startCol, endRow, endCol);
                    }
                    catch (Exception uiEx)
                    {
                        _logger?.Error(uiEx, "🚨 DRAG UI ERROR: Failed to update selection on UI thread");
                    }
                });
            }
            else
            {
                // Fallback to direct call if no dispatcher available
                ClearAllCellSelection();

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

                _logger?.Debug("🎯 SIMPLE DRAG: Selected from [{StartRow},{StartCol}] to [{EndRow},{EndCol}] (direct)", 
                    startRow, startCol, endRow, endCol);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 SIMPLE DRAG ERROR: Failed to update selection");
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
            
            // Add focus event handlers for debugging
            this.GotFocus += AdvancedDataGrid_GotFocus;
            this.LostFocus += AdvancedDataGrid_LostFocus;
            
            // ENHANCED: Add additional keyboard event handlers for maximum coverage
            this.PreviewKeyDown += AdvancedDataGrid_PreviewKeyDown;
            this.KeyUp += AdvancedDataGrid_KeyUp;
            
            // Ensure the control can receive keyboard focus
            this.IsTabStop = true;
            this.UseSystemFocusVisuals = true;
            this.AllowFocusOnInteraction = true;
            
            _logger?.Info("⌨️ KEYBOARD INIT: Enhanced keyboard handlers initialized - KeyDown, PreviewKeyDown, KeyUp, Focus tracking");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 KEYBOARD ERROR: Failed to initialize keyboard handlers");
        }
    }

    /// <summary>
    /// Handle control got focus event for debugging
    /// </summary>
    private void AdvancedDataGrid_GotFocus(object sender, RoutedEventArgs e)
    {
        _logger?.Debug("🎯 CONTROL FOCUS: AdvancedDataGrid got focus - FocusState: {FocusState}", e.OriginalSource?.GetType().Name);
    }

    /// <summary>
    /// Handle control lost focus event for debugging  
    /// </summary>
    private void AdvancedDataGrid_LostFocus(object sender, RoutedEventArgs e)
    {
        _logger?.Debug("🎯 CONTROL FOCUS: AdvancedDataGrid lost focus - FocusState: {FocusState}", e.OriginalSource?.GetType().Name);
    }

    /// <summary>
    /// Handle preview key down - fires before KeyDown
    /// </summary>
    private async void AdvancedDataGrid_PreviewKeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        _logger?.Info("⚡ PREVIEW KEYBOARD: PreviewKeyDown fired - Key: {Key}, Handled: {Handled}, OriginalSource: {OriginalSource}", 
            e.Key, e.Handled, e.OriginalSource?.GetType().Name);
            
        // Check specifically for Ctrl+C in preview
        bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control)
            .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
            
        if (e.Key == Windows.System.VirtualKey.C && isCtrlPressed)
        {
            _logger?.Info("🔥 PREVIEW CTRL+C: Detected Ctrl+C in PreviewKeyDown!");
            
            // Try to handle it directly in preview
            var result = await HandleCtrlCAsync();
            _logger?.Info("🔥 PREVIEW CTRL+C RESULT: Direct handling returned {Result}", result);
            
            // Mark as handled to prevent further processing
            e.Handled = true;
        }
    }

    /// <summary>
    /// Handle key up for debugging
    /// </summary>
    private void AdvancedDataGrid_KeyUp(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        _logger?.Debug("⬆️ KEY UP: Key released - {Key}, OriginalSource: {OriginalSource}", 
            e.Key, e.OriginalSource?.GetType().Name);
    }

    /// <summary>
    /// Handle global keyboard shortcuts according to newProject.md specification
    /// </summary>
    private async void AdvancedDataGrid_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // COMPREHENSIVE LOGGING: Method entry with key information
            _logger?.Info("⌨️ KEYBOARD INPUT: AdvancedDataGrid_KeyDown started - Key: {Key}, Sender: {SenderType}", 
                e.Key, sender?.GetType()?.Name ?? "null");
            _logger?.Debug("⌨️ KEYBOARD INPUT DEBUG: Thread: {ThreadId}, OriginalSource: {OriginalSource}, Handled: {Handled}", 
                Environment.CurrentManagedThreadId, e.OriginalSource?.GetType()?.Name, e.Handled);
            
            bool isCtrlPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Control).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
            
            // COMPREHENSIVE LOGGING: Modifier keys state
            _logger?.Info("⌨️ KEYBOARD MODIFIERS: Key={Key}, Ctrl={Ctrl}, ControlFocusState={ControlFocusState}", 
                e.Key, isCtrlPressed, this.FocusState);
            bool isShiftPressed = Microsoft.UI.Input.InputKeyboardSource.GetKeyStateForCurrentThread(Windows.System.VirtualKey.Shift).HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down);
            
            // DEBUG: Log all key presses to diagnose copy issue
            _logger?.Info("🎹 KEYBOARD DEBUG: Key={Key}, Ctrl={Ctrl}, Shift={Shift}, IsEditing={IsEditing}", 
                e.Key, isCtrlPressed, isShiftPressed, IsCurrentlyEditing());

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
                _logger?.Debug("⌨️ KEYBOARD: Handled shortcut {Key} (Ctrl: {Ctrl}, Shift: {Shift})", 
                    e.Key, isCtrlPressed, isShiftPressed);
            }
            
            // COMPREHENSIVE LOGGING: Method exit with timing
            stopwatch.Stop();
            _logger?.Info("✅ KEYBOARD INPUT: AdvancedDataGrid_KeyDown completed - Key: {Key}, Handled: {Handled}, Time: {ElapsedMs}ms", 
                e.Key, handled, stopwatch.ElapsedMilliseconds);
            _logger?.Debug("✅ KEYBOARD INPUT DEBUG: Final focus state - Control: {ControlFocus}, EditingCell: {IsEditing}", 
                this.FocusState, IsCurrentlyEditing());
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.Error(ex, "🚨 KEYBOARD ERROR: AdvancedDataGrid_KeyDown failed after {ElapsedMs}ms - Key: {Key}", 
                stopwatch.ElapsedMilliseconds, e.Key);
        }
    }

    /// <summary>
    /// Process keyboard shortcuts according to newProject.md specification
    /// </summary>
    private async Task<bool> HandleKeyboardShortcutAsync(Windows.System.VirtualKey key, bool isCtrlPressed, bool isShiftPressed)
    {
        // Check if we're in edit mode for special Tab behavior
        bool isInEditMode = _currentEditingCell != null && _currentEditingCell.IsEditing;
        
        _logger?.Info("⌨️ KEYBOARD DEBUG: Key={Key}, Ctrl={Ctrl}, Shift={Shift}, InEdit={InEdit}", 
            key, isCtrlPressed, isShiftPressed, isInEditMode);

        // EXPLICIT check for Ctrl+C
        if (key == Windows.System.VirtualKey.C && isCtrlPressed && !isShiftPressed)
        {
            _logger?.Info("🔥 EXPLICIT CTRL+C: Detected Ctrl+C combination!");
            return await HandleCtrlCAsync();
        }
            
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
            (Windows.System.VirtualKey.C, true, false, _) => await HandleCtrlCAsync(),
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
            _logger?.Info("⌨️ KEYBOARD: Start edit focused cell (F2)");
            
            // Find the currently focused cell
            var focusedCell = FindFocusedCell();
            if (focusedCell != null && !focusedCell.IsReadOnly)
            {
                StartCellEditing(focusedCell, false); // false = don't clear content
                _logger?.Info("📝 EDIT START: Cell [{Row},{Col}] entered edit mode via F2", 
                    focusedCell.RowIndex, focusedCell.ColumnIndex);
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 EDIT ERROR: Failed to start editing focused cell");
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
            _logger?.Info("⌨️ KEYBOARD: Cancel cell edit (Escape)");
            
            if (_currentEditingCell != null && _originalEditValue != null)
            {
                // Store position before clearing state
                int cellRow = _currentEditingCell.RowIndex;
                int cellColumn = _currentEditingCell.ColumnIndex;
                
                // Restore original value
                _currentEditingCell.DisplayText = _originalEditValue;
                _currentEditingCell.IsEditing = false;
                
                _logger?.Debug("📝 EDIT CANCEL: Cell [{Row},{Col}] editing cancelled, restored to '{Original}'", 
                    _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex, _originalEditValue);
                
                // Clear editing state
                _currentEditingCell = null;
                _originalEditValue = null;
                
                // REFRESH UI after cancel (internal method behavior)
                await TriggerInternalUIRefreshAsync("Cell edit cancelled");
                
                // FOCUS RETENTION: Always restore focus after cancel
                await RestoreFocusAfterRefreshAsync(cellRow, cellColumn);
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 EDIT CANCEL ERROR: Failed to cancel cell editing");
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
            _logger?.Info("⌨️ KEYBOARD: Commit cell edit (Enter), stay on cell: {StayOnCell}", stayOnCell);
            
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                // Store reference and position before clearing state
                var editingCell = _currentEditingCell;
                int cellRow = editingCell.RowIndex;
                int cellColumn = editingCell.ColumnIndex;
                
                // Save the current value to TableCore
                if (_controller != null)
                {
                    await SaveCellValueAsync(editingCell);
                }
                else
                {
                    _logger?.Error("🚨 EDIT COMMIT ERROR: Controller is null, cannot save cell value");
                    return false;
                }
                
                // Exit edit mode
                editingCell.IsEditing = false;
                
                _logger?.Debug("📝 EDIT COMMIT: Cell [{Row},{Col}] editing committed with value '{Value}'", 
                    editingCell.RowIndex, editingCell.ColumnIndex, editingCell.DisplayText);
                
                // Clear editing state
                _currentEditingCell = null;
                _originalEditValue = null;
                
                // REFRESH UI after commit (internal method behavior)
                await TriggerInternalUIRefreshAsync("Cell edit committed");
                
                // FOCUS RETENTION: Restore focus to the committed cell if stayOnCell is true
                if (stayOnCell)
                {
                    await RestoreFocusAfterRefreshAsync(cellRow, cellColumn);
                }
                
                return true;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 EDIT COMMIT ERROR: Failed to commit cell edit");
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
            _logger?.Info("⌨️ KEYBOARD: Insert new line in cell (Shift+Enter)");
            
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
                        
                        _logger?.Debug("📝 MULTILINE: Inserted newline at cursor position {Position} in cell [{Row},{Column}]", 
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
                    
                    _logger?.Debug("📝 MULTILINE FALLBACK: Inserted newline at end in cell [{Row},{Column}]", 
                        _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex);
                    
                    return true;
                }
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 MULTILINE ERROR: Failed to insert new line in cell");
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
            _logger?.Info("⌨️ KEYBOARD: Insert tab in cell (Tab in edit mode)");
            
            if (_currentEditingCell != null && _currentEditingCell.IsEditing)
            {
                // Find the actual TextBox to get cursor position
                var textBox = await FindTextBoxForCellAsync(_currentEditingCell);
                if (textBox != null)
                {
                    // Insert tab character at cursor position
                    int cursorPosition = textBox.SelectionStart;
                    string currentText = textBox.Text ?? string.Empty;
                    string newText = currentText.Insert(cursorPosition, "\t");
                    
                    textBox.Text = newText;
                    textBox.SelectionStart = cursorPosition + 1; // Move cursor after the tab
                    
                    // Update the cell model
                    _currentEditingCell.DisplayText = newText;
                    
                    _logger?.Debug("📋 TAB INSERT: Inserted tab at position {Position} in cell [{Row},{Column}]", 
                        cursorPosition, _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex);
                    
                    return true;
                }
                else
                {
                    // Fallback: Insert tab at the end
                    string currentText = _currentEditingCell.DisplayText ?? string.Empty;
                    string newText = currentText + "\t";
                    _currentEditingCell.DisplayText = newText;
                    
                    _logger?.Debug("📋 TAB INSERT FALLBACK: Inserted tab at end in cell [{Row},{Column}]", 
                        _currentEditingCell.RowIndex, _currentEditingCell.ColumnIndex);
                    
                    return true;
                }
            }
            
            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 TAB INSERT ERROR: Failed to insert tab in cell");
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
            _logger?.Info("⌨️ KEYBOARD: Move to next cell (Tab)");
            
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
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move to next cell");
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
            _logger?.Info("⌨️ KEYBOARD: Move to previous cell (Shift+Tab)");
            
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
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move to previous cell");
            return false;
        }
    }

    /// <summary>
    /// Handle Ctrl+C with explicit debugging
    /// </summary>
    private async Task<bool> HandleCtrlCAsync()
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // COMPREHENSIVE LOGGING: Method entry with context
            _logger?.Info("🔥 CTRL+C PRESSED: Starting enhanced copy operation - Focus: {FocusState}, Thread: {ThreadId}", 
                this.FocusState, Environment.CurrentManagedThreadId);
            
            // ENHANCED FOCUS MANAGEMENT: Use the new aggressive focus method
            await EnsureKeyboardFocusAsync();
            
            // Verify we have selected cells
            var selectedCells = GetAllSelectedCells();
            _logger?.Info("📋 COPY PREPARATION: Found {SelectedCount} selected cells before copy operation", 
                selectedCells.Count);
                
            if (selectedCells.Count == 0)
            {
                _logger?.Warning("📋 COPY WARNING: No cells selected for copy operation");
                return false;
            }
            
            var result = await CopySelectedCellsAsync();
            
            stopwatch.Stop();
            _logger?.Info("✅ CTRL+C COMPLETED: Copy operation returned {Result} in {ElapsedMs}ms", 
                result, stopwatch.ElapsedMilliseconds);
            return result;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.Error(ex, "🚨 CTRL+C ERROR: Copy operation failed after {ElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds);
            return false;
        }
    }

    /// <summary>
    /// Copy selected cells to clipboard (Ctrl+C)
    /// </summary>
    private async Task<bool> CopySelectedCellsAsync()
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // COMPREHENSIVE LOGGING: Method entry with context
            _logger?.Info("📋 COPY OPERATION: CopySelectedCellsAsync started - Focused: [{Row},{Col}], Thread: {ThreadId}", 
                _focusedRowIndex, _focusedColumnIndex, Environment.CurrentManagedThreadId);
            
            if (_uiManager == null) 
            {
                _logger?.Error("📋 COPY ERROR: UIManager is null");
                return false;
            }

            var selectedCells = new List<DataCellModel>();
            var totalCells = 0;
            
            // Collect all selected cells with detailed logging
            foreach (var row in _uiManager.RowsCollection)
            {
                foreach (var cell in row.Cells)
                {
                    totalCells++;
                    
                    // VERBOSE LOGGING: Log every cell state for debugging
                    _logger?.Debug("📋 COPY SCAN: Cell [{Row},{Col}] = '{Value}' (IsSelected: {IsSelected}, IsFocused: {IsFocused})", 
                        cell.RowIndex, cell.ColumnIndex, cell.DisplayText, cell.IsSelected, cell.IsFocused);
                    
                    if (cell.IsSelected)
                    {
                        selectedCells.Add(cell);
                        _logger?.Info("📋 COPY FOUND: Selected cell [{Row},{Col}] = '{Value}'", 
                            cell.RowIndex, cell.ColumnIndex, cell.DisplayText);
                    }
                }
            }

            _logger?.Info("📋 COPY DEBUG: Found {Count} selected cells out of {Total} total cells", selectedCells.Count, totalCells);
            
            // FALLBACK: If no selected cells but we have focused cell, copy just the focused cell
            if (selectedCells.Count == 0 && _focusedCell != null)
            {
                selectedCells.Add(_focusedCell);
                _logger?.Info("📋 COPY FALLBACK: No selected cells, copying focused cell [{Row},{Col}] = '{Value}'", 
                    _focusedCell.RowIndex, _focusedCell.ColumnIndex, _focusedCell.DisplayText);
            }

            if (selectedCells.Count == 0)
            {
                _logger?.Warning("⌨️ COPY: No cells selected");
                return false;
            }

            // Create tab-separated text for clipboard
            var copiedText = await CreateClipboardTextFromSelectedCells(selectedCells);
            
            _logger?.Info("📋 COPY TEXT: Generated clipboard text: '{Text}' (length: {Length})", 
                copiedText?.Length > 100 ? copiedText.Substring(0, 100) + "..." : copiedText, copiedText?.Length ?? 0);
            
            // Copy to system clipboard
            var dataPackage = new Windows.ApplicationModel.DataTransfer.DataPackage();
            dataPackage.SetText(copiedText);
            
            try 
            {
                Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(dataPackage);
                _logger?.Info("✅ CLIPBOARD: Successfully set clipboard content");
            }
            catch (Exception clipEx)
            {
                _logger?.Error(clipEx, "🚨 CLIPBOARD ERROR: Failed to set clipboard content");
                throw;
            }

            // BACKGROUND COLORS: Clear previous copy mode and set new copy mode state
            await ClearAllCopyModeAsync();
            await SetCopyModeForCellsAsync(selectedCells);

            stopwatch.Stop();
            _logger?.Info("✅ COPY OPERATION: CopySelectedCellsAsync completed successfully - Copied {Count} cells in {ElapsedMs}ms", 
                selectedCells.Count, stopwatch.ElapsedMilliseconds);
            return true;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.Error(ex, "🚨 COPY ERROR: CopySelectedCellsAsync failed after {ElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds);
            return false;
        }
    }

    /// <summary>
    /// Clear copy mode from all cells
    /// </summary>
    private async Task ClearAllCopyModeAsync()
    {
        try
        {
            if (_uiManager == null) return;

            // Clear copy mode from all cells and update their visual state
            foreach (var row in _uiManager.RowsCollection)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.IsCopied)
                    {
                        cell.IsCopied = false;
                        UpdateCellSelectionVisuals(cell);
                    }
                }
            }
            
            _logger?.Debug("🎨 COPY MODE: Cleared copy mode from all cells");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 COPY MODE ERROR: Failed to clear copy mode");
        }
    }

    /// <summary>
    /// Set copy mode for specific cells and update their visual state
    /// </summary>
    private async Task SetCopyModeForCellsAsync(List<DataCellModel> cells)
    {
        try
        {
            foreach (var cell in cells)
            {
                cell.IsCopied = true;
                UpdateCellSelectionVisuals(cell);
            }
            
            _logger?.Debug("🎨 COPY MODE: Set copy mode for {Count} cells", cells.Count);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 COPY MODE ERROR: Failed to set copy mode for cells");
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
                _logger?.Warning("⌨️ PASTE: No text content in clipboard");
                return false;
            }

            var clipboardText = await dataPackageView.GetTextAsync();
            if (string.IsNullOrEmpty(clipboardText))
            {
                _logger?.Warning("⌨️ PASTE: Clipboard text is empty");
                return false;
            }

            // BACKGROUND COLORS: Clear copy mode when pasting
            await ClearAllCopyModeAsync();

            // Find focused cell as paste target
            var focusedCell = GetFocusedCell();
            if (focusedCell == null)
            {
                _logger?.Warning("⌨️ PASTE: No focused cell for paste target");
                return false;
            }

            // Parse clipboard text and paste
            await PasteClipboardTextToCells(clipboardText, focusedCell);

            _logger?.Info("⌨️ PASTE: Pasted clipboard content starting at [{Row},{Col}]", 
                focusedCell.RowIndex, focusedCell.ColumnIndex);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 PASTE ERROR: Failed to paste from clipboard");
            return false;
        }
    }

    /// <summary>
    /// Cut selected cells to clipboard (Ctrl+X)
    /// </summary>
    private async Task<bool> CutSelectedCellsAsync()
    {
        _logger?.Info("⌨️ KEYBOARD: Cut selected cells (Ctrl+X)");
        // TODO: Implement clipboard cut logic
        return true;
    }

    /// <summary>
    /// Select all cells (Ctrl+A)
    /// </summary>
    private async Task<bool> SelectAllCellsAsync()
    {
        _logger?.Info("⌨️ KEYBOARD: Select all cells (Ctrl+A)");
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
            _logger?.Info("⌨️ KEYBOARD: Smart delete (Delete)");
            
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
                    
                    _logger?.Debug("🗑️ SMART DELETE: Cleared cell [{Row},{Col}]", 
                        cell.RowIndex, cell.ColumnIndex);
                }
            }
            
            _logger?.Info("✅ SMART DELETE: Cleared {Count} cells", cellsToDelete.Count);
            
            // REFRESH UI after delete (internal method behavior)
            if (cellsToDelete.Count > 0)
            {
                await TriggerInternalUIRefreshAsync($"Smart delete of {cellsToDelete.Count} cells");
            }
            
            return cellsToDelete.Count > 0;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 DELETE ERROR: Failed to smart delete cells");
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
            _logger?.Info("⌨️ KEYBOARD: Force delete row (Ctrl+Delete)");
            
            if (_focusedCell == null) 
            {
                _logger?.Warning("🚨 CTRL+DELETE: No focused cell available, _focusedCell is null");
                return false;
            }
            
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
            
            _logger?.Info("✅ ROW DELETE: Deleted row {Row}, remaining rows: {Total}", 
                rowToDelete, TableCore.ActualRowCount);
                
            return true;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW DELETE ERROR: Failed to delete row");
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
            _logger?.Info("⌨️ KEYBOARD: Insert row above (Insert)");
            
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
            
            _logger?.Info("✅ ROW INSERT: Inserted empty row at position {Position}, total rows: {Total}", 
                insertPosition, TableCore.ActualRowCount);
                
            return true;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW INSERT ERROR: Failed to insert row");
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
            _logger?.Info("⌨️ KEYBOARD: Move to first cell (Ctrl+Home)");
            return await MoveFocusToAsync(0, 0);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move to first cell");
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
            _logger?.Info("⌨️ KEYBOARD: Move to last data cell (Ctrl+End)");
            
            // Find the last row with data
            int lastDataRow = await TableCore.GetLastDataRowAsync();
            if (lastDataRow < 0) lastDataRow = 0; // If no data, go to first row
            
            int lastColumn = Math.Max(0, ColumnCount - 1);
            
            return await MoveFocusToAsync(lastDataRow, lastColumn);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move to last data cell");
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
            _logger?.Debug("⌨️ KEYBOARD: Move up (Arrow Up)");
            
            int newRow = Math.Max(0, _focusedRowIndex - 1);
            return await MoveFocusToAsync(newRow, _focusedColumnIndex);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move up");
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
            _logger?.Debug("⌨️ KEYBOARD: Move down (Arrow Down)");
            
            int newRow = Math.Min(ActualRowCount - 1, _focusedRowIndex + 1);
            return await MoveFocusToAsync(newRow, _focusedColumnIndex);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move down");
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
            _logger?.Debug("⌨️ KEYBOARD: Move left (Arrow Left)");
            
            int newColumn = Math.Max(0, _focusedColumnIndex - 1);
            return await MoveFocusToAsync(_focusedRowIndex, newColumn);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move left");
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
            _logger?.Debug("⌨️ KEYBOARD: Move right (Arrow Right)");
            
            int newColumn = Math.Min(ColumnCount - 1, _focusedColumnIndex + 1);
            return await MoveFocusToAsync(_focusedRowIndex, newColumn);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move right");
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
                _logger?.Warning("🚨 NAVIGATION: Invalid cell position [{Row},{Column}] - bounds: [{MaxRow},{MaxColumn}]", 
                    row, column, ActualRowCount - 1, ColumnCount - 1);
                return false;
            }

            // Update focus tracking
            _focusedRowIndex = row;
            _focusedColumnIndex = column;

            // Clear previous focus visual but keep selection
            if (_focusedCell != null)
            {
                _focusedCell.IsFocused = false;
                // CRITICAL: Update visuals to preserve selection coloring
                UpdateCellSelectionVisuals(_focusedCell);
            }

            // Get new focused cell
            if (row < _uiManager.RowsCollection.Count && column < _uiManager.RowsCollection[row].Cells.Count)
            {
                _focusedCell = _uiManager.RowsCollection[row].Cells[column];
                _focusedCell.IsFocused = true;
                
                _logger?.Debug("🎯 FOCUS: Moved to cell [{Row},{Column}] - '{DisplayText}'", 
                    row, column, _focusedCell.DisplayText);
                
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 NAVIGATION ERROR: Failed to move focus to [{Row},{Column}]", row, column);
            return false;
        }
    }

    /// <summary>
    /// Enhanced focus management for keyboard events (especially Ctrl+C)
    /// </summary>
    private async Task EnsureKeyboardFocusAsync()
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            _logger?.Info("🎯 ENHANCED FOCUS: Starting aggressive focus acquisition for keyboard events");
            
            // Step 1: Check current focus state
            bool initialFocusState = this.FocusState != FocusState.Unfocused;
            _logger?.Debug("🎯 ENHANCED FOCUS DEBUG: Initial focus state - HasFocus: {HasFocus}, FocusState: {FocusState}", 
                initialFocusState, this.FocusState);
            
            // Step 2: Try multiple focus approaches with small delays
            bool[] focusResults = new bool[4];
            
            focusResults[0] = this.Focus(FocusState.Programmatic);
            await Task.Delay(5); // Small delay to allow focus to settle
            
            focusResults[1] = this.Focus(FocusState.Keyboard);
            await Task.Delay(5);
            
            focusResults[2] = this.Focus(FocusState.Pointer);
            await Task.Delay(5);
            
            // Step 3: Force focus on container elements if needed
            if (this.FocusState == FocusState.Unfocused)
            {
                // Try to focus on the main container (this control)
                this.IsTabStop = true;
                this.UseSystemFocusVisuals = true;
                focusResults[3] = this.Focus(FocusState.Keyboard);
                await Task.Delay(10);
            }
            
            // Step 4: Final verification
            bool finalFocusState = this.FocusState != FocusState.Unfocused;
            
            _logger?.Info("✅ ENHANCED FOCUS: Completed - Initial: {Initial}, Final: {Final}, Results: [{R0},{R1},{R2},{R3}]", 
                initialFocusState, finalFocusState, focusResults[0], focusResults[1], focusResults[2], focusResults[3]);
            
            stopwatch.Stop();
            _logger?.Debug("✅ ENHANCED FOCUS DEBUG: Completed in {ElapsedMs}ms - FocusState: {FinalState}", 
                stopwatch.ElapsedMilliseconds, this.FocusState);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.Error(ex, "🚨 ENHANCED FOCUS ERROR: Failed after {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
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
            
            _logger?.Debug("🎯 RANGE SELECT: Selected range ({StartRow},{StartCol}) to ({EndRow},{EndCol}) - {Count} cells",
                startRow, startCol, endRow, endCol, (endRow - startRow + 1) * (endCol - startCol + 1));
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 RANGE SELECT ERROR: Failed to select cell range");
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
            
            _logger?.Debug("🧹 CLEAR SELECT: Cleared all cell selections with visual updates");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 CLEAR SELECT ERROR: Failed to clear selections");
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
            
            _logger?.Debug("🔍 TEXTBOX: Looking for editing TextBox for cell [{Row},{Column}]", 
                cell.RowIndex, cell.ColumnIndex);
            
            // This would need to be implemented based on your actual XAML structure
            // For now, we'll simulate this by working with the cell's DisplayText directly
            return null; // TODO: Implement actual TextBox finding logic
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 TEXTBOX ERROR: Failed to find editing TextBox for cell [{Row},{Column}]", 
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
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            _logger?.Info("📋 CLIPBOARD TEXT: CreateClipboardTextFromSelectedCells started - {Count} cells", selectedCells.Count);
            
            if (selectedCells.Count == 0) 
            {
                _logger?.Warning("📋 CLIPBOARD TEXT: No cells provided for clipboard text creation");
                return string.Empty;
            }

            // Group cells by row and sort
            var cellsByRow = selectedCells
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .ToList();

            _logger?.Info("📋 CLIPBOARD TEXT: Grouped into {RowCount} rows", cellsByRow.Count);

            var lines = new List<string>();
            
            foreach (var rowGroup in cellsByRow)
            {
                var cellsInRow = rowGroup.OrderBy(c => c.ColumnIndex).ToList();
                var cellValues = cellsInRow.Select(c => c.DisplayText ?? string.Empty).ToList();
                var rowText = string.Join("\t", cellValues);
                lines.Add(rowText);
                
                _logger?.Debug("📋 CLIPBOARD ROW: Row {Row} = '{Text}' (from {CellCount} cells)", 
                    rowGroup.Key, rowText.Length > 50 ? rowText.Substring(0, 50) + "..." : rowText, cellsInRow.Count);
            }

            var result = string.Join("\r\n", lines);
            
            stopwatch.Stop();
            _logger?.Info("✅ CLIPBOARD TEXT: Created clipboard text - {Length} characters, {Lines} lines in {ElapsedMs}ms", 
                result.Length, lines.Count, stopwatch.ElapsedMilliseconds);
                
            return result;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.Error(ex, "🚨 CLIPBOARD TEXT ERROR: Failed to create clipboard text after {ElapsedMs}ms", 
                stopwatch.ElapsedMilliseconds);
            return string.Empty;
        }
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

            _logger?.Info("📋 PASTE: Pasted {Lines} lines x {Cols} columns to grid", 
                lines.Length, lines.Length > 0 ? lines[0].Split('\t').Length : 0);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 PASTE ERROR: Failed to parse and paste clipboard text");
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

    /// <summary>
    /// Get all currently selected cells
    /// </summary>
    private List<DataCellModel> GetAllSelectedCells()
    {
        var selectedCells = new List<DataCellModel>();
        if (_uiManager == null) return selectedCells;
        
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
        return selectedCells;
    }

    #endregion

    #endregion

}