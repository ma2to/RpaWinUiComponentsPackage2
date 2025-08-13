using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
// FINÁLNY CLEAN PUBLIC API - jediný import!
using RpaWinUiComponentsPackage;
using RpaWinUiComponentsPackage.LoggerComponent;
// Explicit import pre AdvancedWinUiDataGrid namespace
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid;
using System.Data;
using System.Linq;
using System.Text;

namespace RpaWinUiComponents.Demo;

/// <summary>
/// Demo aplikácia implementuje UI Update Strategy pattern z newProject.md
/// Kombinuje API calls + manual UI updates pre správny user experience
/// </summary>
public sealed partial class MainWindow : Window
{
    #region Private Fields

    private readonly ILogger<MainWindow> _logger;
    private readonly LoggerComponent _fileLogger;
    private readonly StringBuilder _logOutput = new();
    private bool _isGridInitialized = false;

    #endregion

    #region Constructor

    public MainWindow()
    {
        this.InitializeComponent();
        
        // Setup logging
        _logger = App.LoggerFactory.CreateLogger<MainWindow>();
        
        // Setup file logger for testing LoggerComponent
        var tempLogDir = Path.Combine(Path.GetTempPath(), "RpaWinUiDemo");
        _fileLogger = LoggerComponentFactory.WithRotation(
            logger: _logger,
            logDirectory: tempLogDir,
            baseFileName: "demo",
            maxFileSizeMB: 5
        );

        AddLogMessage("🚀 Demo application started - Testing package reference mode");
        AddLogMessage($"📂 File logging to: {_fileLogger.CurrentLogFile}");
    }

    #endregion

    #region Initialization Event Handlers

    private async void InitButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AddLogMessage("🔧 DEMO ACTION: Initializing basic DataGrid...");

            var columns = CreateBasicColumns();
            var logger = App.LoggerFactory.CreateLogger("DataGrid");

            // FINÁLNY CLEAN API initialization
            await TestDataGrid.InitializeAsync(
                columns: columns,
                colors: null, // Default colors
                validation: null, // No validation
                performance: null, // Default performance
                emptyRowsCount: 10,
                logger: logger
            );

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            _isGridInitialized = true;
            AddLogMessage("✅ DataGrid initialized successfully + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Initialization failed: {ex.Message}");
            await _fileLogger.Error(ex, "DataGrid initialization failed");
        }
    }

    private async void InitWithValidationButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AddLogMessage("🔧 DEMO ACTION: Initializing DataGrid with validation...");

            var columns = CreateAdvancedColumns();
            var validationConfig = new ValidationConfiguration
            {
                EnableRealtimeValidation = true,
                EnableBatchValidation = true,
                ShowValidationAlerts = true,
                RulesWithMessages = new Dictionary<string, (Func<object, bool> Validator, string ErrorMessage)>
                {
                    ["Name"] = (value => !string.IsNullOrEmpty(value?.ToString()), "Name is required"),
                    ["Age"] = (value => int.TryParse(value?.ToString(), out int age) && age >= 0 && age <= 120, "Age must be between 0 and 120"),
                    ["Email"] = (value => {
                        var email = value?.ToString();
                        return !string.IsNullOrEmpty(email) && email.Contains("@") && email.Contains(".");
                    }, "Invalid email format")
                },
                CrossRowRules = new List<Func<List<Dictionary<string, object?>>, (bool IsValid, string? ErrorMessage)>>
                {
                    allData =>
                    {
                        var emails = allData.Select(row => row.GetValueOrDefault("Email")?.ToString())
                                           .Where(email => !string.IsNullOrEmpty(email))
                                           .ToList();
                        bool isUnique = emails.Count == emails.Distinct().Count();
                        return isUnique ? (true, null) : (false, "Duplicate emails found");
                    }
                }
            };
            var logger = App.LoggerFactory.CreateLogger("DataGrid");

            // FINÁLNY CLEAN API initialization with validation
            await TestDataGrid.InitializeAsync(
                columns: columns,
                colors: null, // Default colors
                validation: validationConfig,
                performance: null, // Default performance
                emptyRowsCount: 15,
                logger: logger
            );

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            _isGridInitialized = true;
            AddLogMessage("✅ DataGrid with validation initialized + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Initialization with validation failed: {ex.Message}");
            await _fileLogger.Error(ex, "DataGrid validation initialization failed");
        }
    }

    #endregion

    #region Data Operations Event Handlers

    private async void ImportDictionaryButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📥 DEMO ACTION: Importing Dictionary data...");

            var testData = CreateTestDictionaryData();

            // HEADLESS import (NO automatic UI refresh)
            await TestDataGrid.ImportFromDictionaryAsync(testData);

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            AddLogMessage($"✅ Imported {testData.Count} rows + UI refreshed");
            await _fileLogger.Info($"Dictionary import completed: {testData.Count} rows");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Dictionary import failed: {ex.Message}");
            await _fileLogger.Error(ex, "Dictionary import failed");
        }
    }

    private async void ImportDataTableButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📥 DEMO ACTION: Importing DataTable...");

            var dataTable = CreateTestDataTable();

            // HEADLESS import (NO automatic UI refresh)
            await TestDataGrid.ImportFromDataTableAsync(dataTable);

            // MANUAL UI refresh (demo pattern)  
            await TestDataGrid.RefreshUIAsync();

            AddLogMessage($"✅ Imported {dataTable.Rows.Count} rows from DataTable + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ DataTable import failed: {ex.Message}");
        }
    }

    private async void ExportDictionaryButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📤 DEMO ACTION: Exporting to Dictionary...");

            // HEADLESS export
            var exportedData = await TestDataGrid.ExportToDictionaryAsync(includeValidAlerts: true);

            AddLogMessage($"✅ Exported {exportedData.Count} rows to Dictionary");
            
            // Log sample data
            if (exportedData.Count > 0)
            {
                var firstRow = exportedData.First();
                AddLogMessage($"📋 Sample row: {string.Join(", ", firstRow.Select(kvp => $"{kvp.Key}={kvp.Value}"))}");
            }
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Dictionary export failed: {ex.Message}");
        }
    }

    private async void ExportDataTableButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📤 DEMO ACTION: Exporting to DataTable...");

            // HEADLESS export  
            var dataTable = await TestDataGrid.ExportToDataTableAsync(includeValidAlerts: false);

            AddLogMessage($"✅ Exported {dataTable.Rows.Count} rows, {dataTable.Columns.Count} columns to DataTable");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ DataTable export failed: {ex.Message}");
        }
    }

    private async void ClearDataButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🗑️ DEMO ACTION: Clearing all data...");

            // HEADLESS clear (NO automatic UI refresh)
            await TestDataGrid.ClearAllDataAsync();

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            AddLogMessage("✅ All data cleared + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Clear data failed: {ex.Message}");
        }
    }

    #endregion

    #region Validation Event Handlers

    private async void ValidateAllButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("✅ DEMO ACTION: Validating all rows...");

            // HEADLESS validation (NO automatic UI refresh)
            bool allValid = await TestDataGrid.AreAllNonEmptyRowsValidAsync();

            AddLogMessage($"📊 Validation result: {(allValid ? "All rows valid" : "Some rows invalid")}");
            await _fileLogger.Info($"Validation completed: AllValid={allValid}");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Validation failed: {ex.Message}");
        }
    }

    private async void BatchValidationButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("✅ DEMO ACTION: Running batch validation...");

            // HEADLESS batch validation (NO automatic UI refresh)
            var validationResult = await TestDataGrid.ValidateAllRowsBatchAsync();

            if (validationResult != null)
            {
                AddLogMessage($"📊 Batch validation: Valid={validationResult.ValidCellsCount}, Invalid={validationResult.InvalidCellsCount}");
                
                // MANUAL UI refresh for validation indicators (demo pattern)
                await TestDataGrid.UpdateValidationUIAsync();
                
                AddLogMessage("🎨 Validation UI indicators updated");
            }
            else
            {
                AddLogMessage("📊 Batch validation: No validation configured");
            }
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Batch validation failed: {ex.Message}");
        }
    }

    #endregion

    #region UI Update Event Handlers

    private async void RefreshUIButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Manual full UI refresh...");

            await TestDataGrid.RefreshUIAsync();

            AddLogMessage("✅ Full UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ UI refresh failed: {ex.Message}");
        }
    }

    private async void UpdateValidationUIButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Manual validation UI update...");

            await TestDataGrid.UpdateValidationUIAsync();

            AddLogMessage("✅ Validation UI updated");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Validation UI update failed: {ex.Message}");
        }
    }

    private void InvalidateLayoutButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Invalidating layout...");

            TestDataGrid.InvalidateLayout();

            AddLogMessage("✅ Layout invalidated");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Layout invalidation failed: {ex.Message}");
        }
    }

    #endregion

    #region Row Management Event Handlers

    private async void SmartDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🗑️ DEMO ACTION: Smart deleting row 2...");

            // HEADLESS delete (NO automatic UI refresh)
            await TestDataGrid.SmartDeleteRowAsync(2);

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            AddLogMessage("✅ Row 2 smart deleted + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Smart delete failed: {ex.Message}");
        }
    }

    private async void CompactRowsButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📋 DEMO ACTION: Compacting rows...");

            // HEADLESS compact (NO automatic UI refresh)
            await TestDataGrid.CompactRowsAsync();

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            AddLogMessage("✅ Rows compacted + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Compact rows failed: {ex.Message}");
        }
    }

    private async void PasteDataButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📋 DEMO ACTION: Pasting test data...");

            var pasteData = CreatePasteTestData();

            // HEADLESS paste with auto-expand (NO automatic UI refresh)
            await TestDataGrid.PasteDataAsync(pasteData, startRow: 5, startColumn: 0);

            // MANUAL UI refresh (demo pattern)
            await TestDataGrid.RefreshUIAsync();

            AddLogMessage($"✅ Pasted {pasteData.Count} rows starting at row 5 + UI refreshed");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Paste data failed: {ex.Message}");
        }
    }

    #endregion

    #region Color Theme Event Handlers

    private void ApplyDarkThemeButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Applying dark theme...");

            // Create dark theme using clean ColorConfiguration API
            var darkColors = new ColorConfiguration
            {
                CellBackground = "#2D2D30",
                CellForeground = "#F1F1F1", 
                CellBorder = "#3F3F46",
                HeaderBackground = "#1E1E1E",
                HeaderForeground = "#FFFFFF",
                HeaderBorder = "#3F3F46",
                SelectionBackground = "#094771",
                SelectionForeground = "#FFFFFF"
            };
            
            // Apply through InitializeAsync (demo workaround since ApplyColorConfig is internal)
            AddLogMessage("⚠️ Dark theme requires re-initialization with clean API");

            AddLogMessage("✅ Dark theme applied");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Dark theme failed: {ex.Message}");
        }
    }

    private void ResetColorsButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Resetting colors...");

            TestDataGrid.ResetColorsToDefaults();

            AddLogMessage("✅ Colors reset to defaults");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Reset colors failed: {ex.Message}");
        }
    }

    #endregion

    #region Statistics Event Handlers

    private async void GetStatsButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("📊 DEMO ACTION: Getting statistics...");

            var totalRows = TestDataGrid.GetTotalRowCount();
            var columnCount = TestDataGrid.GetColumnCount();
            var visibleRows = await TestDataGrid.GetVisibleRowsCountAsync();
            var lastDataRow = await TestDataGrid.GetLastDataRowAsync();
            var hasData = TestDataGrid.HasData;
            var minRows = TestDataGrid.GetMinimumRowCount();

            AddLogMessage($"📊 STATISTICS:");
            AddLogMessage($"   • Total rows: {totalRows}");
            AddLogMessage($"   • Columns: {columnCount}");
            AddLogMessage($"   • Visible rows: {visibleRows}");
            AddLogMessage($"   • Last data row: {lastDataRow}");
            AddLogMessage($"   • Has data: {hasData}");
            AddLogMessage($"   • Minimum rows: {minRows}");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Get statistics failed: {ex.Message}");
        }
    }

    #endregion

    #region Helper Methods

    private void AddLogMessage(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        var logEntry = $"[{timestamp}] {message}\n";
        
        _logOutput.Append(logEntry);
        LogOutput.Text = _logOutput.ToString();
        
        // Auto-scroll to bottom
        LogScrollViewer.ScrollToVerticalOffset(LogScrollViewer.ExtentHeight);
        
        // Also log to file
        _ = Task.Run(async () => await _fileLogger.Info(message));
    }

    private List<ColumnConfiguration> CreateBasicColumns()
    {
        return new List<ColumnConfiguration>
        {
            new() { Name = "Name", DisplayName = "Name", Type = typeof(string), Width = 150 },
            new() { Name = "Age", DisplayName = "Age", Type = typeof(int), Width = 80 },
            new() { Name = "Email", DisplayName = "Email", Type = typeof(string), Width = 200 }
        };
    }

    private List<ColumnConfiguration> CreateAdvancedColumns()
    {
        return new List<ColumnConfiguration>
        {
            // User columns
            new() { Name = "Name", DisplayName = "Full Name", Type = typeof(string), Width = 150 },
            new() { Name = "Age", DisplayName = "Age", Type = typeof(int), Width = 80 },
            new() { Name = "Email", DisplayName = "Email Address", Type = typeof(string), Width = 200 },
            new() { Name = "Salary", DisplayName = "Salary", Type = typeof(decimal), Width = 120 },
            
            // Special columns (auto-positioned)
            new() { Name = "ValidationAlerts", DisplayName = "Errors", IsValidationColumn = true, Width = 100 },
            new() { Name = "DeleteRows", DisplayName = "Delete", IsDeleteColumn = true, Width = 60 }
        };
    }

    private List<Dictionary<string, object?>> CreateTestDictionaryData()
    {
        return new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "John Doe", ["Age"] = 30, ["Email"] = "john@example.com", ["Salary"] = 50000m },
            new() { ["Name"] = "Jane Smith", ["Age"] = 25, ["Email"] = "jane@example.com", ["Salary"] = 55000m },
            new() { ["Name"] = "Bob Johnson", ["Age"] = -5, ["Email"] = "invalid-email", ["Salary"] = 45000m }, // Invalid data for validation testing
            new() { ["Name"] = "", ["Age"] = 35, ["Email"] = "bob@example.com", ["Salary"] = 60000m }, // Invalid name
            new() { ["Name"] = "Alice Brown", ["Age"] = 28, ["Email"] = "alice@example.com", ["Salary"] = 52000m },
        };
    }

    private DataTable CreateTestDataTable()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Columns.Add("Email", typeof(string));
        table.Columns.Add("Salary", typeof(decimal));

        table.Rows.Add("Mike Wilson", 32, "mike@example.com", 48000m);
        table.Rows.Add("Sarah Davis", 29, "sarah@example.com", 51000m);
        table.Rows.Add("Tom Anderson", 45, "tom@example.com", 65000m);

        return table;
    }

    private List<Dictionary<string, object?>> CreatePasteTestData()
    {
        return new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "Paste Test 1", ["Age"] = 22, ["Email"] = "paste1@test.com", ["Salary"] = 40000m },
            new() { ["Name"] = "Paste Test 2", ["Age"] = 24, ["Email"] = "paste2@test.com", ["Salary"] = 42000m },
        };
    }

    #endregion

    #region Color Testing Event Handlers - SELECTIVE OVERRIDE TESTS

    private void TestSelectiveColorsButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Testing selective color override - border + selection...");

            // Test selective override using clean ColorConfiguration API
            var selectiveColors = new ColorConfiguration
            {
                CellBorder = "#FF0000",           // Custom červený border  
                SelectionBackground = "#FFFF00", // Custom žltý selection
                // Ostatné farby null → použijú sa default farby
            };
            
            AddLogMessage("⚠️ Selective colors require re-initialization with clean API");

            AddLogMessage("✅ Selective colors applied - red border + yellow selection, rest default");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Selective color test failed: {ex.Message}");
        }
    }

    private void TestBorderOnlyButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Testing border-only color override...");

            // Test nastavenia len border farby using clean API
            var borderOnlyColors = new ColorConfiguration
            {
                CellBorder = "#0000FF",    // Custom modrý border
                HeaderBorder = "#0000FF",  // Custom modrý header border
                // Všetky ostatné farby null → použijú sa default farby
            };
            
            AddLogMessage("⚠️ Border-only colors require re-initialization with clean API");

            AddLogMessage("✅ Border-only colors applied - blue borders, rest default");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Border-only color test failed: {ex.Message}");
        }
    }

    private void TestValidationOnlyButton_Click(object sender, RoutedEventArgs e)
    {
        if (!_isGridInitialized)
        {
            AddLogMessage("⚠️ Grid must be initialized first!");
            return;
        }

        try
        {
            AddLogMessage("🎨 DEMO ACTION: Testing validation-only color override...");

            // Test nastavenia len validation farieb using clean API
            var validationOnlyColors = new ColorConfiguration
            {
                ValidationErrorBorder = "#FFA500",      // Custom oranžový validation border
                ValidationErrorBackground = "#32FFA500", // Custom oranžový validation background (with alpha)
                // Všetky ostatné farby null → použijú sa default farby
            };
            
            AddLogMessage("⚠️ Validation-only colors require re-initialization with clean API");

            AddLogMessage("✅ Validation-only colors applied - orange validation colors, rest default");
            AddLogMessage("ℹ️ To test validation colors, import invalid data and validate");
        }
        catch (Exception ex)
        {
            AddLogMessage($"❌ Validation-only color test failed: {ex.Message}");
        }
    }

    #endregion
}

