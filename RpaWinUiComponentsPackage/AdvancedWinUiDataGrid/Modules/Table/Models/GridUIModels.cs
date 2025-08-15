using Microsoft.UI;
using Microsoft.UI.Xaml.Media;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.ObjectModel;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;

/// <summary>
/// UI Model pre header cell v DataGrid s proper data binding support
/// </summary>
public class HeaderCellModel : INotifyPropertyChanged
{
    private string _displayName = string.Empty;
    private string _columnName = string.Empty;
    private double _width = 100;
    private bool _isSortable = true;
    private bool _isFilterable = true;
    private SolidColorBrush _backgroundBrush = new SolidColorBrush(Colors.LightGray);
    private SolidColorBrush _foregroundBrush = new SolidColorBrush(Colors.Black);
    private SolidColorBrush _borderBrush = new SolidColorBrush(Colors.Gray);

    public string DisplayName
    {
        get => _displayName;
        set => SetProperty(ref _displayName, value);
    }

    public string ColumnName
    {
        get => _columnName;
        set => SetProperty(ref _columnName, value);
    }

    public double Width
    {
        get => _width;
        set => SetProperty(ref _width, value);
    }

    public bool IsSortable
    {
        get => _isSortable;
        set => SetProperty(ref _isSortable, value);
    }

    public bool IsFilterable
    {
        get => _isFilterable;
        set => SetProperty(ref _isFilterable, value);
    }

    public SolidColorBrush BackgroundBrush
    {
        get => _backgroundBrush;
        set => SetProperty(ref _backgroundBrush, value);
    }

    public SolidColorBrush ForegroundBrush
    {
        get => _foregroundBrush;
        set => SetProperty(ref _foregroundBrush, value);
    }

    public SolidColorBrush BorderBrush
    {
        get => _borderBrush;
        set => SetProperty(ref _borderBrush, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}

/// <summary>
/// UI Model pre data cell v DataGrid s proper data binding support
/// </summary>
public class DataCellModel : INotifyPropertyChanged
{
    private object? _value;
    private string _displayText = string.Empty;
    private int _rowIndex;
    private int _columnIndex;
    private string _columnName = string.Empty;
    private bool _isReadOnly = false;
    private bool _isValid = true;
    private bool _isEditing = false;
    private string? _validationError;
    private double _width = 100; // Width for proper header/cell alignment
    private SolidColorBrush _backgroundBrush = new SolidColorBrush(Colors.White);
    private SolidColorBrush _foregroundBrush = new SolidColorBrush(Colors.Black);
    private SolidColorBrush _borderBrush = new SolidColorBrush(Colors.LightGray);

    public object? Value
    {
        get => _value;
        set => SetProperty(ref _value, value);
    }

    public string DisplayText
    {
        get => _displayText;
        set => SetProperty(ref _displayText, value);
    }

    public int RowIndex
    {
        get => _rowIndex;
        set => SetProperty(ref _rowIndex, value);
    }

    public int ColumnIndex
    {
        get => _columnIndex;
        set => SetProperty(ref _columnIndex, value);
    }

    public string ColumnName
    {
        get => _columnName;
        set => SetProperty(ref _columnName, value);
    }

    public bool IsReadOnly
    {
        get => _isReadOnly;
        set => SetProperty(ref _isReadOnly, value);
    }

    public bool IsValid
    {
        get => _isValid;
        set => SetProperty(ref _isValid, value);
    }

    public bool IsEditing
    {
        get => _isEditing;
        set => SetProperty(ref _isEditing, value);
    }

    public string? ValidationError
    {
        get => _validationError;
        set => SetProperty(ref _validationError, value);
    }

    public double Width
    {
        get => _width;
        set => SetProperty(ref _width, value);
    }

    public SolidColorBrush BackgroundBrush
    {
        get => _backgroundBrush;
        set => SetProperty(ref _backgroundBrush, value);
    }

    public SolidColorBrush ForegroundBrush
    {
        get => _foregroundBrush;
        set => SetProperty(ref _foregroundBrush, value);
    }

    public SolidColorBrush BorderBrush
    {
        get => _borderBrush;
        set => SetProperty(ref _borderBrush, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}

/// <summary>
/// UI Model pre riadok v DataGrid s proper data binding support
/// </summary>
public class DataRowModel : INotifyPropertyChanged
{
    private int _rowIndex;
    private bool _isSelected = false;
    private bool _isValid = true;
    private bool _isEmpty = true;
    private SolidColorBrush _backgroundBrush = new SolidColorBrush(Colors.White);

    public int RowIndex
    {
        get => _rowIndex;
        set => SetProperty(ref _rowIndex, value);
    }

    public ObservableCollection<DataCellModel> Cells { get; set; } = new();

    public bool IsSelected
    {
        get => _isSelected;
        set => SetProperty(ref _isSelected, value);
    }

    public bool IsValid
    {
        get => _isValid;
        set => SetProperty(ref _isValid, value);
    }

    public bool IsEmpty
    {
        get => _isEmpty;
        set => SetProperty(ref _isEmpty, value);
    }

    public SolidColorBrush BackgroundBrush
    {
        get => _backgroundBrush;
        set => SetProperty(ref _backgroundBrush, value);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}

/// <summary>
/// Kompletný UI Model pre DataGrid
/// </summary>
public class DataGridUIModel
{
    public List<HeaderCellModel> Headers { get; set; } = new();
    public List<DataRowModel> Rows { get; set; } = new();
    public int TotalRowCount { get; set; }
    public int VisibleRowCount { get; set; }
    public bool IsLoading { get; set; } = false;
    public string StatusText { get; set; } = "Ready";
}