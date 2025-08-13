using Microsoft.Extensions.Logging.Abstractions;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Performance.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Validation.Models;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Services;

/// <summary>
/// Headless core pre AdvancedDataGrid - HYBRID model implementation
/// Základ systému - funguje bez UI pre script scenarios
/// Implementuje "Always All Validations" a "Intelligent Row Management" stratégiu
/// </summary>
public class DynamicTableCore
{
    #region Private Fields

    /// <summary>
    /// Row-based storage pre fast bulk operations (Import/Export/Search performance)
    /// </summary>
    private readonly List<DataRow> _rows = new();

    /// <summary>
    /// Column definitions
    /// </summary>
    private readonly List<GridColumnDefinition> _columns = new();

    /// <summary>
    /// Minimálny počet riadkov (definovaný z aplikácie)
    /// </summary>
    private int _minimumRowCount = 15;

    /// <summary>
    /// Validation configuration (implementované v aplikácii)
    /// </summary>
    private IValidationConfiguration? _validationConfig;

    /// <summary>
    /// Logger (nullable - funguje bez loggera)
    /// </summary>
    private readonly Microsoft.Extensions.Logging.ILogger? _logger;

    /// <summary>
    /// Throttling configuration
    /// </summary>
    private GridThrottlingConfig _throttlingConfig = GridThrottlingConfig.Default;

    /// <summary>
    /// Je core inicializovaný
    /// </summary>
    private bool _isInitialized = false;

    #endregion

    #region Properties

    /// <summary>
    /// Skutočný počet riadkov v gridu (intelligent row management)
    /// </summary>
    public int ActualRowCount => _rows.Count;

    /// <summary>
    /// Minimálny počet riadkov
    /// </summary>
    public int MinimumRowCount => _minimumRowCount;

    /// <summary>
    /// Počet stĺpcov
    /// </summary>
    public int ColumnCount => _columns.Count;

    /// <summary>
    /// Získa zoznam názvov všetkých stĺpcov (v poradí)
    /// </summary>
    public List<string> GetColumnNames()
    {
        return _columns.Select(c => c.Name).ToList();
    }

    /// <summary>
    /// Je core inicializovaný
    /// </summary>
    public bool IsInitialized => _isInitialized;

    /// <summary>
    /// Má dataset nejaké dáta
    /// </summary>
    public bool HasData => _rows.Any(row => !row.IsEmpty);

    /// <summary>
    /// Počet riadkov s dátami
    /// </summary>
    public int DataRowCount => _rows.Count(row => !row.IsEmpty);

    #endregion

    #region Constructor

    public DynamicTableCore(Microsoft.Extensions.Logging.ILogger? logger = null)
    {
        _logger = logger;
    }

    #endregion

    #region Initialization

    /// <summary>
    /// Inicializuje table core s column definitions
    /// </summary>
    public async Task InitializeAsync(
        List<GridColumnDefinition> columns,
        IValidationConfiguration? validationConfig = null,
        GridThrottlingConfig? throttlingConfig = null,
        int emptyRowsCount = 15,
        Microsoft.Extensions.Logging.ILogger? logger = null)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // ROZSIAHLE LOGOVANIE - method entry s parametrami
            _logger?.LogMethodEntry(nameof(InitializeAsync), 
                columns?.Count, validationConfig?.IsValidationEnabled, emptyRowsCount,
                throttlingConfig?.EnableAggressiveMemoryManagement, throttlingConfig?.EnableBackgroundProcessing);
            
            // Log column details
            _logger?.LogDataDetails("INPUT_COLUMNS", columns, columns?.Count);
            
            // Validation column definitions
            var validationStopwatch = System.Diagnostics.Stopwatch.StartNew();
            await ValidateColumnDefinitionsAsync(columns);
            validationStopwatch.Stop();
            _logger?.LogPerformance("COLUMN_VALIDATION", validationStopwatch.Elapsed, columns?.Count);

            // Set configuration
            _validationConfig = validationConfig;
            _throttlingConfig = throttlingConfig ?? GridThrottlingConfig.Default;
            _minimumRowCount = emptyRowsCount;

            // Clear existing data
            _rows.Clear();
            _columns.Clear();

            // Apply automatic special column positioning
            var arrangementStopwatch = System.Diagnostics.Stopwatch.StartNew();
            var arrangedColumns = ApplySpecialColumnPositioning(columns);
            _columns.AddRange(arrangedColumns);
            arrangementStopwatch.Stop();
            _logger?.LogPerformance("COLUMN_ARRANGEMENT", arrangementStopwatch.Elapsed, arrangedColumns.Count);

            // Create minimum rows + 1 empty row at end
            var rowCreationStopwatch = System.Diagnostics.Stopwatch.StartNew();
            await CreateMinimumRowsAsync();
            rowCreationStopwatch.Stop();
            _logger?.LogPerformance("ROW_CREATION", rowCreationStopwatch.Elapsed, _rows.Count);

            _isInitialized = true;

            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(InitializeAsync), stopwatch.Elapsed, "CORE_INITIALIZED");
            _logger?.LogPerformance("CORE_INITIALIZATION", stopwatch.Elapsed, _columns.Count * _rows.Count);
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger?.LogMethodExit(nameof(InitializeAsync), stopwatch.Elapsed, "FAILED");
            _logger?.Error(ex, "🚨 INIT ERROR: DynamicTableCore initialization failed after {Time}ms", stopwatch.ElapsedMilliseconds);
            throw;
        }
    }

    #endregion

    #region Core Data Operations (Headless API)

    /// <summary>
    /// Získa hodnotu bunky
    /// </summary>
    public async Task<object?> GetCellValueAsync(int row, int column)
    {
        ValidatePosition(row, column);
        
        var dataRow = _rows[row];
        var columnName = _columns[column].Name;
        
        return dataRow.GetCellValue(columnName);
    }

    /// <summary>
    /// Nastaví hodnotu bunky s intelligent row expansion
    /// </summary>
    public async Task SetCellValueAsync(int row, int column, object? value)
    {
        try
        {
            ValidatePosition(row, column);

            var dataRow = _rows[row];
            var columnName = _columns[column].Name;
            var oldValue = dataRow.GetCellValue(columnName);

            // Set value
            dataRow.SetCellValue(columnName, value);

            _logger?.Info("📊 DATA CONTEXT: Cell value changed - Row: {Row}, Column: '{Column}', Old: '{OldValue}', New: '{NewValue}'", 
                                   row, columnName, oldValue, value);

            // Auto-expand: ak píšem do posledného prázdneho riadku → pridaj nový prázdny
            if (row == _rows.Count - 1 && !dataRow.IsEmpty)
            {
                await AddEmptyRowAsync();
                _logger?.Info("📊 ROW MANAGEMENT: Auto-expanded table - added empty row, Total rows: {TotalRows}", _rows.Count);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 DATA ERROR: SetCellValueAsync failed - Row: {Row}, Column: {Column}, Value: '{Value}'", 
                             row, column, value);
            throw;
        }
    }

    /// <summary>
    /// Získa celý riadok ako Dictionary
    /// </summary>
    public async Task<Dictionary<string, object?>> GetRowDataAsync(int rowIndex)
    {
        ValidateRowIndex(rowIndex);
        return _rows[rowIndex].GetAllData();
    }

    /// <summary>
    /// Nastaví celý riadok z Dictionary
    /// </summary>
    public async Task SetRowDataAsync(int rowIndex, Dictionary<string, object?> data)
    {
        ValidateRowIndex(rowIndex);
        _rows[rowIndex].SetAllData(data);

        // Auto-expand logic
        var dataRow = _rows[rowIndex];
        if (rowIndex == _rows.Count - 1 && !dataRow.IsEmpty)
        {
            await AddEmptyRowAsync();
        }
    }

    #endregion

    #region Intelligent Row Management

    /// <summary>
    /// Je riadok prázdny (všetky bunky null/empty)
    /// </summary>
    public bool IsRowEmpty(int rowIndex)
    {
        ValidateRowIndex(rowIndex);
        return _rows[rowIndex].IsEmpty;
    }

    /// <summary>
    /// Smart delete riadku - content vs. whole row based on row count
    /// </summary>
    public async Task SmartDeleteRowAsync(int rowIndex)
    {
        ValidateRowIndex(rowIndex);

        // Smart delete logic
        if (_rows.Count > _minimumRowCount + 1)
        {
            // Máme viac ako minimum + 1 → zmaž celý riadok
            await DeleteCompleteRowAsync(rowIndex);
            _logger?.Info("👤 USER ACTION: Complete row deleted - Row: {Row}", rowIndex);
        }
        else
        {
            // Máme minimum alebo menej → vyčisti len obsah, zachovaj štruktúru
            await ClearRowContentAsync(rowIndex);
            _logger?.Info("👤 USER ACTION: Row content cleared - Row: {Row}", rowIndex);
        }
    }

    /// <summary>
    /// Force delete row regardless of minimum count
    /// </summary>
    public async Task ForceDeleteRowAsync(int rowIndex)
    {
        ValidateRowIndex(rowIndex);
        await DeleteCompleteRowAsync(rowIndex);
        _logger?.Info("👤 USER ACTION: Force delete - Row: {Row}", rowIndex);
    }

    /// <summary>
    /// Check if row can be deleted (based on minimum count)
    /// </summary>
    public bool CanDeleteRow(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _rows.Count)
            return false;
            
        return _rows.Count > _minimumRowCount + 1;
    }

    /// <summary>
    /// Paste data s automatic row expansion
    /// </summary>
    public async Task PasteDataAsync(List<Dictionary<string, object?>> pasteData, int startRow, int startColumn)
    {
        ValidateRowIndex(startRow);

        // Automatic row expansion pre paste operations
        int requiredRows = startRow + pasteData.Count;
        if (requiredRows > _rows.Count - 1) // -1 pre prázdny riadok na konci
        {
            await ExpandRowsToCountAsync(requiredRows + 1); // +1 pre nový prázdny riadok
        }

        // Vlož dáta
        for (int i = 0; i < pasteData.Count; i++)
        {
            await SetRowDataAsync(startRow + i, pasteData[i]);
        }

        _logger?.Info("👤 USER ACTION: Data pasted - StartRow: {StartRow}, Count: {Count}", 
                               startRow, pasteData.Count);
    }

    /// <summary>
    /// Index posledného riadku obsahujúceho dáta (-1 ak všetky prázdne)
    /// </summary>
    public async Task<int> GetLastDataRowAsync()
    {
        for (int i = _rows.Count - 1; i >= 0; i--)
        {
            if (!_rows[i].IsEmpty)
                return i;
        }
        return -1; // Všetky riadky prázdne
    }

    /// <summary>
    /// Odstráni prázdne medzery medzi riadkami s dátami
    /// </summary>
    public async Task CompactRowsAsync()
    {
        var dataRows = new List<DataRow>();
        
        // Collect all non-empty rows
        foreach (var row in _rows)
        {
            if (!row.IsEmpty)
            {
                dataRows.Add(row);
            }
        }

        // Clear and rebuild
        _rows.Clear();
        
        // Add compacted data rows
        for (int i = 0; i < dataRows.Count; i++)
        {
            dataRows[i].RowIndex = i;
            _rows.Add(dataRows[i]);
        }

        // Ensure minimum row count + 1 empty row
        await EnsureMinimumRowsAsync();

        _logger?.Info("✅ OPERATION SUCCESS: Rows compacted - Data rows: {DataRows}, Total rows: {TotalRows}", 
                               dataRows.Count, _rows.Count);
    }

    #endregion

    #region Import/Export Operations

    /// <summary>
    /// Import z Dictionary list s intelligent row expansion
    /// </summary>
    public async Task ImportFromDictionaryAsync(
        List<Dictionary<string, object?>> data,
        Dictionary<int, bool>? checkboxStates = null,
        int? startRow = null,
        bool insertMode = false,
        TimeSpan? timeout = null,
        IProgress<ValidationProgress>? validationProgress = null)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        timeout ??= TimeSpan.FromMinutes(1);

        try
        {
            _logger?.Info("🚀 OPERATION START: ImportFromDictionaryAsync - Rows: {Count}, StartRow: {StartRow}, InsertMode: {InsertMode}", 
                                   data.Count, startRow, insertMode);

            int targetRow = startRow ?? 0;

            // Expand rows if needed
            int requiredRows = targetRow + data.Count;
            if (requiredRows > _rows.Count - 1)
            {
                await ExpandRowsToCountAsync(requiredRows + 1);
            }

            // Import data
            for (int i = 0; i < data.Count; i++)
            {
                if (sw.Elapsed > timeout)
                    throw new TimeoutException($"Import operation exceeded timeout of {timeout}");

                await SetRowDataAsync(targetRow + i, data[i]);

                // Set checkbox states if provided
                if (checkboxStates?.TryGetValue(i, out bool isChecked) == true)
                {
                    // TODO: Set checkbox state (when checkbox column is implemented)
                }

                // Report progress
                validationProgress?.Report(new ValidationProgress
                {
                    ProcessedItems = i + 1,
                    TotalItems = data.Count,
                    CurrentOperation = "Importing data"
                });
            }

            // Batch validation (Always All Validations strategy)
            if (_validationConfig?.EnableBatchValidation == true)
            {
                await ValidateAllRowsBatchAsync();
            }

            _logger?.Info("✅ OPERATION SUCCESS: ImportFromDictionaryAsync - Imported: {Rows} rows, Time: {ElapsedMs}ms", 
                                   data.Count, sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 DATA ERROR: Import failed - Format: Dictionary, Size: {Count} rows", data.Count);
            throw;
        }
    }

    /// <summary>
    /// Export do Dictionary list
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ExportToDictionaryAsync(
        bool includeValidationAlerts = false,
        bool removeAfter = false,
        TimeSpan? timeout = null,
        IProgress<ExportProgress>? exportProgress = null)
    {
        timeout ??= TimeSpan.FromMinutes(1);
        var sw = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            _logger?.Info("🚀 OPERATION START: ExportToDictionaryAsync - Rows: {Count}, IncludeValidation: {IncludeValidation}, RemoveAfter: {RemoveAfter}", 
                                   _rows.Count, includeValidationAlerts, removeAfter);

            var result = new List<Dictionary<string, object?>>();
            var exportColumns = GetExportColumns(includeValidationAlerts);

            for (int i = 0; i < _rows.Count; i++)
            {
                if (sw.Elapsed > timeout)
                    throw new TimeoutException($"Export operation exceeded timeout of {timeout}");

                var row = _rows[i];
                if (row.IsEmpty) continue; // Skip empty rows

                var exportRow = new Dictionary<string, object?>();
                
                foreach (var column in exportColumns)
                {
                    exportRow[column.Name] = row.GetCellValue(column.Name);
                }

                result.Add(exportRow);

                // Report progress
                exportProgress?.Report(new ExportProgress
                {
                    ProcessedItems = i + 1,
                    TotalItems = _rows.Count,
                    CurrentOperation = "Exporting data"
                });
            }

            // Remove exported data if requested using smart delete logic
            if (removeAfter && result.Count > 0)
            {
                _logger?.Info("🗑️ OPERATION: RemoveAfter enabled - smart deleting exported data");
                
                // Smart delete all exported rows (those with data)
                var rowsToDelete = new List<int>();
                for (int i = 0; i < _rows.Count; i++)
                {
                    if (!_rows[i].IsEmpty)
                    {
                        rowsToDelete.Add(i);
                    }
                }
                
                // Process deletions in descending order to avoid index issues
                foreach (var rowIndex in rowsToDelete.OrderByDescending(x => x))
                {
                    await SmartDeleteRowAsync(rowIndex);
                }
            }

            _logger?.Info("✅ OPERATION SUCCESS: ExportToDictionaryAsync - Exported: {Rows} rows, Time: {ElapsedMs}ms, RemoveAfter: {RemoveAfter}", 
                                   result.Count, sw.ElapsedMilliseconds, removeAfter);

            return result;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 DATA ERROR: Export failed - Format: Dictionary");
            throw;
        }
    }

    #endregion

    #region Validation Operations (Always All Validations Strategy)

    /// <summary>
    /// Validuje VŠETKY riadky v dataset, nie len zobrazené
    /// </summary>
    public async Task<bool> AreAllNonEmptyRowsValidAsync()
    {
        if (_validationConfig == null || !_validationConfig.IsValidationEnabled)
            return true;

        try
        {
            var allDataRows = await GetAllDataRowsAsync();
            var validationRules = _validationConfig.GetValidationRules();
            var crossRowRules = _validationConfig.GetCrossRowValidationRules();

            // Cell-level validation
            foreach (var row in allDataRows)
            {
                foreach (var column in _columns.Where(c => !c.IsSpecialColumn))
                {
                    var value = row.GetCellValue(column.Name);
                    var rules = validationRules.GetRules(column.Name);

                    foreach (var rule in rules.Where(r => r.IsEnabled))
                    {
                        var result = rule.Validate(value);
                        if (!result.IsValid)
                        {
                            _logger?.Info("📊 VALIDATION: Cell validation failed - Row: {Row}, Column: {Column}, Error: {Error}", 
                                             row.RowIndex, column.Name, result.ErrorMessage);
                            return false;
                        }
                    }
                }
            }

            // Cross-row validation
            var allRowData = allDataRows.Select(r => r.GetAllData()).ToList();
            foreach (var rule in crossRowRules.Where(r => r.IsEnabled))
            {
                var result = rule.Validate(allRowData);
                if (!result.IsValid)
                {
                    _logger?.Info("📊 VALIDATION: Cross-row validation failed - Rule: {Rule}, Error: {Error}", 
                                     rule.Name, result.GlobalErrorMessage);
                    return false;
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 VALIDATION ERROR: AreAllNonEmptyRowsValidAsync failed");
            return false;
        }
    }

    /// <summary>
    /// Batch validation všetkých riadkov
    /// </summary>
    public async Task<BatchValidationResult?> ValidateAllRowsBatchAsync(CancellationToken cancellationToken = default)
    {
        if (_validationConfig == null || !_validationConfig.IsValidationEnabled)
            return null;

        try
        {
            _logger?.Info("🚀 OPERATION START: ValidateAllRowsBatchAsync - Rows: {Count}", _rows.Count);

            var result = new BatchValidationResult();
            var allDataRows = await GetAllDataRowsAsync();
            var validationRules = _validationConfig.GetValidationRules();

            foreach (var row in allDataRows)
            {
                if (cancellationToken.IsCancellationRequested)
                    break;

                foreach (var column in _columns.Where(c => !c.IsSpecialColumn))
                {
                    var value = row.GetCellValue(column.Name);
                    var rules = validationRules.GetRules(column.Name);
                    var cellUIState = row.GetCellUIState(column.Name);

                    // Reset previous validation state
                    cellUIState.ResetValidation();

                    foreach (var rule in rules.Where(r => r.IsEnabled))
                    {
                        var validationResult = rule.Validate(value);
                        if (!validationResult.IsValid)
                        {
                            cellUIState.HasValidationError = true;
                            cellUIState.ValidationErrorMessage = validationResult.ErrorMessage;
                            result.AddError(row.RowIndex, column.Name, validationResult.ErrorMessage);
                            
                            _logger?.Info("🚨 VALIDATION ERROR: Cell validation failed - Row: {Row}, Column: '{Column}', Rule: '{Rule}', Value: '{Value}', Error: '{Error}'", 
                                                   row.RowIndex, column.Name, rule.Name, value, validationResult.ErrorMessage);
                            break; // First error wins
                        }
                        else
                        {
                            _logger?.Info("✅ VALIDATION SUCCESS: Cell validation passed - Row: {Row}, Column: '{Column}', Rule: '{Rule}', Value: '{Value}'", 
                                                   row.RowIndex, column.Name, rule.Name, value);
                        }
                    }
                }
            }

            // Cross-row validation
            var crossRowRules = _validationConfig.GetCrossRowValidationRules();
            var allRowData = allDataRows.Select(r => r.GetAllData()).ToList();

            foreach (var rule in crossRowRules.Where(r => r.IsEnabled))
            {
                var crossRowResult = rule.Validate(allRowData);
                if (!crossRowResult.IsValid)
                {
                    result.AddCrossRowErrors(crossRowResult);
                }
            }

            _logger?.Info("✅ OPERATION SUCCESS: Validation completed - Valid: {Valid}, Invalid: {Invalid}", 
                                   result.ValidCellsCount, result.InvalidCellsCount);

            return result;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 VALIDATION ERROR: Batch validation failed");
            throw;
        }
    }

    #endregion

    #region Private Helper Methods

    /// <summary>
    /// Validuje column definitions
    /// </summary>
    private async Task ValidateColumnDefinitionsAsync(List<GridColumnDefinition> columns)
    {
        if (columns == null || columns.Count == 0)
            throw new ArgumentException("At least one column definition is required");

        foreach (var column in columns)
        {
            if (!column.IsValid(out string? errorMessage))
            {
                throw new ArgumentException($"Invalid column definition '{column.Name}': {errorMessage}");
            }
        }
    }

    /// <summary>
    /// Aplikuje automatic special column positioning
    /// </summary>
    private List<GridColumnDefinition> ApplySpecialColumnPositioning(List<GridColumnDefinition> originalColumns)
    {
        var result = new List<GridColumnDefinition>();
        var userColumns = originalColumns.Where(c => !c.IsSpecialColumn).ToList();
        var checkboxColumn = originalColumns.FirstOrDefault(c => c.IsCheckBoxColumn);
        var validationColumn = originalColumns.FirstOrDefault(c => c.IsValidationAlertsColumn);
        var deleteColumn = originalColumns.FirstOrDefault(c => c.IsDeleteRowColumn);

        // 1. CheckBox (if enabled): FIRST position
        if (checkboxColumn != null)
            result.Add(checkboxColumn);

        // 2. User-defined columns: MIDDLE positions
        result.AddRange(userColumns);

        // 3. ValidationAlerts: SECOND-TO-LAST position (if no DeleteRows) or SECOND-TO-LAST
        if (validationColumn != null)
            result.Add(validationColumn);

        // 4. DeleteRows (if enabled): LAST position
        if (deleteColumn != null)
            result.Add(deleteColumn);

        _logger?.Info("📊 COLUMN POSITIONING: Final order - CheckBox: {HasCheckbox}, User: {UserCount}, Validation: {HasValidation}, Delete: {HasDelete}",
                         checkboxColumn != null, userColumns.Count, validationColumn != null, deleteColumn != null);

        return result;
    }

    /// <summary>
    /// Vytvorí minimálny počet riadkov + 1 prázdny na konci
    /// </summary>
    private async Task CreateMinimumRowsAsync()
    {
        for (int i = 0; i < _minimumRowCount + 1; i++) // +1 pre prázdny riadok na konci
        {
            _rows.Add(new DataRow(i));
        }
    }

    /// <summary>
    /// Zabezpečí minimálny počet riadkov
    /// </summary>
    private async Task EnsureMinimumRowsAsync()
    {
        while (_rows.Count < _minimumRowCount + 1)
        {
            _rows.Add(new DataRow(_rows.Count));
        }
    }

    /// <summary>
    /// Pridá prázdny riadok na koniec
    /// </summary>
    private async Task AddEmptyRowAsync()
    {
        _rows.Add(new DataRow(_rows.Count));
    }

    /// <summary>
    /// Rozšíri počet riadkov na target count
    /// </summary>
    private async Task ExpandRowsToCountAsync(int targetCount)
    {
        while (_rows.Count < targetCount)
        {
            _rows.Add(new DataRow(_rows.Count));
        }
    }

    /// <summary>
    /// Zmaže kompletný riadok
    /// </summary>
    private async Task DeleteCompleteRowAsync(int rowIndex)
    {
        _rows.RemoveAt(rowIndex);
        
        // Re-index remaining rows
        for (int i = rowIndex; i < _rows.Count; i++)
        {
            _rows[i].RowIndex = i;
        }

        // Ensure minimum rows
        await EnsureMinimumRowsAsync();
    }

    /// <summary>
    /// Vyčistí obsah riadku (zachová štruktúru)
    /// </summary>
    private async Task ClearRowContentAsync(int rowIndex)
    {
        _rows[rowIndex].ClearData();
    }

    /// <summary>
    /// Získa všetky data rows (všetky, nie len viewport)
    /// </summary>
    private async Task<List<DataRow>> GetAllDataRowsAsync()
    {
        return _rows.Where(row => !row.IsEmpty).ToList();
    }

    /// <summary>
    /// Získa columns pre export (filter special columns based on includeValidationAlerts)
    /// </summary>
    private List<GridColumnDefinition> GetExportColumns(bool includeValidationAlerts)
    {
        return _columns.Where(c => 
            !c.IsCheckBoxColumn && 
            !c.IsDeleteRowColumn && 
            (!c.IsValidationAlertsColumn || includeValidationAlerts))
        .ToList();
    }

    /// <summary>
    /// Validuje pozíciu bunky
    /// </summary>
    private void ValidatePosition(int row, int column)
    {
        ValidateRowIndex(row);
        if (column < 0 || column >= _columns.Count)
            throw new ArgumentOutOfRangeException(nameof(column), "Column index out of range");
    }

    /// <summary>
    /// Validuje index riadku
    /// </summary>
    private void ValidateRowIndex(int row)
    {
        if (row < 0 || row >= _rows.Count)
            throw new ArgumentOutOfRangeException(nameof(row), "Row index out of range");
    }

    #endregion
}

/// <summary>
/// Progress tracking pre validation
/// </summary>
public class ValidationProgress
{
    public int ProcessedItems { get; set; }
    public int TotalItems { get; set; }
    public string CurrentOperation { get; set; } = string.Empty;
    public double PercentComplete => TotalItems > 0 ? (double)ProcessedItems / TotalItems * 100 : 0;
}

/// <summary>
/// Progress tracking pre export
/// </summary>
public class ExportProgress
{
    public int ProcessedItems { get; set; }
    public int TotalItems { get; set; }
    public string CurrentOperation { get; set; } = string.Empty;
    public double PercentComplete => TotalItems > 0 ? (double)ProcessedItems / TotalItems * 100 : 0;
}

/// <summary>
/// Výsledok batch validation
/// </summary>
public class BatchValidationResult
{
    private readonly Dictionary<string, List<string>> _cellErrors = new();
    private readonly Dictionary<int, string> _rowErrors = new();

    public int ValidCellsCount { get; private set; }
    public int InvalidCellsCount => _cellErrors.Values.Sum(errors => errors.Count);

    public void AddError(int rowIndex, string columnName, string errorMessage)
    {
        var key = $"{rowIndex}:{columnName}";
        if (!_cellErrors.ContainsKey(key))
            _cellErrors[key] = new List<string>();
        
        _cellErrors[key].Add(errorMessage);
    }

    public void AddCrossRowErrors(CrossRowValidationResult crossRowResult)
    {
        foreach (var kvp in crossRowResult.RowErrors)
        {
            _rowErrors[kvp.Key] = kvp.Value;
        }
    }

    public bool HasErrors => InvalidCellsCount > 0 || _rowErrors.Count > 0;
}