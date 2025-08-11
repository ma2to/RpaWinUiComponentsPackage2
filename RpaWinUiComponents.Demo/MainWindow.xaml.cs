using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Configuration;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Validation;
using RpaWinUiComponentsPackage.LoggerComponent;
using System.Data;
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

            // HEADLESS initialization
            await TestDataGrid.InitializeAsync(
                columns: columns,
                validationConfig: null,
                throttlingConfig: GridThrottlingConfig.Default,
                emptyRowsCount: 10,
                colorConfig: DataGridColorConfig.Default,
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
            var validationConfig = new DemoValidationConfiguration();
            var logger = App.LoggerFactory.CreateLogger("DataGrid");

            // HEADLESS initialization with validation
            await TestDataGrid.InitializeAsync(
                columns: columns,
                validationConfig: validationConfig,
                throttlingConfig: GridThrottlingConfig.Default,
                emptyRowsCount: 15,
                colorConfig: DataGridColorConfig.Default,
                logger: logger,
                enableBatchValidation: true
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
            TestDataGrid.SmartDeleteRowAsync(2);

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

            TestDataGrid.ApplyColorConfig(DataGridColorConfig.Dark);

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

    private List<GridColumnDefinition> CreateBasicColumns()
    {
        return new List<GridColumnDefinition>
        {
            new() { Name = "Name", DisplayName = "Name", DataType = typeof(string), Width = 150 },
            new() { Name = "Age", DisplayName = "Age", DataType = typeof(int), Width = 80 },
            new() { Name = "Email", DisplayName = "Email", DataType = typeof(string), Width = 200 },
        };
    }

    private List<GridColumnDefinition> CreateAdvancedColumns()
    {
        return new List<GridColumnDefinition>
        {
            // User columns
            new() { Name = "Name", DisplayName = "Full Name", DataType = typeof(string), Width = 150 },
            new() { Name = "Age", DisplayName = "Age", DataType = typeof(int), Width = 80 },
            new() { Name = "Email", DisplayName = "Email Address", DataType = typeof(string), Width = 200 },
            new() { Name = "Salary", DisplayName = "Salary", DataType = typeof(decimal), Width = 120 },
            
            // Special columns (auto-positioned)
            new() { Name = "ValidationAlerts", DisplayName = "Errors", IsValidationAlertsColumn = true, Width = 100 },
            new() { Name = "DeleteRows", DisplayName = "Delete", IsDeleteRowColumn = true, Width = 60 }
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
}

/// <summary>
/// Demo validation configuration - implementované v aplikácii, NIE v balíku
/// </summary>
public class DemoValidationConfiguration : IValidationConfiguration
{
    public bool IsValidationEnabled => true;
    public bool EnableRealtimeValidation => true;
    public bool EnableBatchValidation => true;

    public ValidationRuleSet GetValidationRules()
    {
        var ruleSet = new ValidationRuleSet();

        // Name validation
        ruleSet.AddRule("Name", new ValidationRule
        {
            Name = "NameRequired",
            Validator = value => !string.IsNullOrEmpty(value?.ToString()),
            ErrorMessage = "Name is required"
        });

        // Age validation
        ruleSet.AddRule("Age", new ValidationRule
        {
            Name = "ValidAge",
            Validator = value => int.TryParse(value?.ToString(), out int age) && age >= 0 && age <= 120,
            ErrorMessage = "Age must be between 0 and 120"
        });

        // Email validation
        ruleSet.AddRule("Email", new ValidationRule
        {
            Name = "EmailFormat",
            Validator = value => 
            {
                var email = value?.ToString();
                return !string.IsNullOrEmpty(email) && email.Contains("@") && email.Contains(".");
            },
            ErrorMessage = "Invalid email format"
        });

        return ruleSet;
    }

    public List<CrossRowValidationRule> GetCrossRowValidationRules()
    {
        return new List<CrossRowValidationRule>
        {
            new CrossRowValidationRule
            {
                Name = "UniqueEmails",
                Validator = allData =>
                {
                    var emails = allData.Select(row => row.GetValueOrDefault("Email")?.ToString())
                                       .Where(email => !string.IsNullOrEmpty(email))
                                       .ToList();
                    
                    bool isUnique = emails.Count == emails.Distinct().Count();
                    return isUnique ? 
                        CrossRowValidationResult.Success() : 
                        CrossRowValidationResult.Error("Duplicate emails found");
                }
            }
        };
    }
}