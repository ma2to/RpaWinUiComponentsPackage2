namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Core;

/// <summary>
/// HYBRID model - kombinuje Dictionary pre bulk operácie a CellUIState pre individual cell management
/// Best of both worlds: Row-based storage (fast bulk) + Cell objects (UI state)
/// </summary>
public class DataRow
{
    /// <summary>
    /// Row-based storage pre fast bulk operations (Import/Export/Search performance)
    /// </summary>
    private readonly Dictionary<string, object?> _data;

    /// <summary>
    /// Cell UI states pre individual cell management (Selection/Validation/Styling flexibility)
    /// </summary>
    private readonly Dictionary<string, CellUIState> _cellStates;

    /// <summary>
    /// Row index v gridu
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// Je riadok prázdny (všetky bunky null/empty)
    /// </summary>
    public bool IsEmpty => _data.Values.All(value => value == null || string.IsNullOrEmpty(value.ToString()));

    /// <summary>
    /// Má riadok nejaké zmeny
    /// </summary>
    public bool HasChanges => _cellStates.Values.Any(state => state.HasChanges);

    /// <summary>
    /// Má riadok validation errors
    /// </summary>
    public bool HasValidationErrors => _cellStates.Values.Any(state => state.HasValidationError);

    /// <summary>
    /// Je riadok označený (aspoň jedna bunka selected)
    /// </summary>
    public bool IsSelected => _cellStates.Values.Any(state => state.IsSelected);

    /// <summary>
    /// Počet buniek s dátami v riadku
    /// </summary>
    public int NonEmptyCellsCount => _data.Values.Count(value => value != null && !string.IsNullOrEmpty(value.ToString()));

    public DataRow(int rowIndex)
    {
        RowIndex = rowIndex;
        _data = new Dictionary<string, object?>();
        _cellStates = new Dictionary<string, CellUIState>();
    }

    /// <summary>
    /// Získa hodnotu bunky
    /// </summary>
    public object? GetCellValue(string columnName)
    {
        return _data.TryGetValue(columnName, out var value) ? value : null;
    }

    /// <summary>
    /// Nastaví hodnotu bunky a označí ju ako zmenenú
    /// </summary>
    public void SetCellValue(string columnName, object? value)
    {
        var oldValue = GetCellValue(columnName);
        if (!Equals(oldValue, value))
        {
            _data[columnName] = value;
            EnsureCellUIState(columnName).HasChanges = true;
        }
    }

    /// <summary>
    /// Získa UI stav bunky
    /// </summary>
    public CellUIState GetCellUIState(string columnName)
    {
        return EnsureCellUIState(columnName);
    }

    /// <summary>
    /// Zabezpečí existenciu UI stavu pre bunku
    /// </summary>
    private CellUIState EnsureCellUIState(string columnName)
    {
        if (!_cellStates.TryGetValue(columnName, out var state))
        {
            state = new CellUIState();
            _cellStates[columnName] = state;
        }
        return state;
    }

    /// <summary>
    /// Vyčistí všetky dáta v riadku (zachováva UI state)
    /// </summary>
    public void ClearData()
    {
        _data.Clear();
        // UI state zostáva pre zachovanie selection, atd.
    }

    /// <summary>
    /// Vyčistí všetky dáta aj UI stavy
    /// </summary>
    public void ClearAll()
    {
        _data.Clear();
        _cellStates.Clear();
    }

    /// <summary>
    /// Reset zmien (dirty flags)
    /// </summary>
    public void ResetChanges()
    {
        foreach (var state in _cellStates.Values)
        {
            state.HasChanges = false;
        }
    }

    /// <summary>
    /// Reset validation errors
    /// </summary>
    public void ResetValidationErrors()
    {
        foreach (var state in _cellStates.Values)
        {
            state.ResetValidation();
        }
    }

    /// <summary>
    /// Reset selection stavov
    /// </summary>
    public void ResetSelection()
    {
        foreach (var state in _cellStates.Values)
        {
            state.ResetSelection();
        }
    }

    /// <summary>
    /// Získa všetky dáta ako Dictionary (pre bulk operations)
    /// </summary>
    public Dictionary<string, object?> GetAllData()
    {
        return new Dictionary<string, object?>(_data);
    }

    /// <summary>
    /// Nastaví všetky dáta z Dictionary (pre bulk operations)
    /// </summary>
    public void SetAllData(Dictionary<string, object?> data)
    {
        _data.Clear();
        foreach (var kvp in data)
        {
            _data[kvp.Key] = kvp.Value;
        }
    }

    /// <summary>
    /// Získa zoznam stĺpcov s dátami
    /// </summary>
    public IEnumerable<string> GetColumnsWithData()
    {
        return _data.Where(kvp => kvp.Value != null && !string.IsNullOrEmpty(kvp.Value.ToString()))
                   .Select(kvp => kvp.Key);
    }

    /// <summary>
    /// Získa zoznam stĺpcov s validation errors
    /// </summary>
    public IEnumerable<string> GetColumnsWithValidationErrors()
    {
        return _cellStates.Where(kvp => kvp.Value.HasValidationError)
                         .Select(kvp => kvp.Key);
    }
}