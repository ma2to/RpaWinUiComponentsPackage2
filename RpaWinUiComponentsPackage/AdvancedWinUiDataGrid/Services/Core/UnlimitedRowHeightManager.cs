using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Text;
using Windows.Foundation;
using Windows.UI.Text;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Configuration;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Services.Core;

/// <summary>
/// Unlimited Row Height Manager - automatické prispôsobovanie výšky riadkov obsahu
/// Všetky riadky majú jednotnú výšku, ale ak obsah presahuje → výška sa automaticky zvýši pre všetky riadky
/// NO LIMITS: výška môže byť neobmedzená pre zobrazenie celého obsahu
/// </summary>
public class UnlimitedRowHeightManager
{
    #region Private Fields

    /// <summary>
    /// Default base row height (32px)
    /// </summary>
    private const double DefaultBaseRowHeight = 32.0;

    /// <summary>
    /// Minimum row height (16px)
    /// </summary>
    private const double MinRowHeight = 16.0;

    /// <summary>
    /// Current unified row height (všetky riadky majú rovnakú výšku)
    /// </summary>
    private double _currentUnifiedRowHeight = DefaultBaseRowHeight;

    /// <summary>
    /// Base row height configuration
    /// </summary>
    private double _baseRowHeight = DefaultBaseRowHeight;

    /// <summary>
    /// Column width information pre text measurement
    /// </summary>
    private Dictionary<string, double> _columnWidths = new();

    /// <summary>
    /// Font information pre text measurement
    /// </summary>
    private FontInfo _fontInfo = new();

    /// <summary>
    /// Logger (nullable)
    /// </summary>
    private readonly Microsoft.Extensions.Logging.ILogger? _logger;

    /// <summary>
    /// Is height calculation enabled
    /// </summary>
    private bool _isEnabled = true;

    #endregion

    #region Properties

    /// <summary>
    /// Aktuálna unified row height (všetky riadky)
    /// </summary>
    public double CurrentUnifiedRowHeight => _currentUnifiedRowHeight;

    /// <summary>
    /// Base row height (default pre prázdne riadky)
    /// </summary>
    public double BaseRowHeight 
    { 
        get => _baseRowHeight; 
        set 
        { 
            if (value >= MinRowHeight)
            {
                _baseRowHeight = value;
                _logger?.Info("📐 ROW HEIGHT: Base height updated to {Height}px", value);
            }
        } 
    }

    /// <summary>
    /// Je unlimited row height enabled
    /// </summary>
    public bool IsEnabled 
    { 
        get => _isEnabled; 
        set 
        { 
            _isEnabled = value; 
            _logger?.Info("📐 ROW HEIGHT: Unlimited height {Status}", value ? "ENABLED" : "DISABLED");
        } 
    }

    #endregion

    #region Constructor

    /// <summary>
    /// Konštruktor s optional logger
    /// </summary>
    public UnlimitedRowHeightManager(Microsoft.Extensions.Logging.ILogger? logger = null)
    {
        _logger = logger;
        _currentUnifiedRowHeight = _baseRowHeight;

        _logger?.Info("📐 ROW HEIGHT MANAGER: Initialized - Base height: {Height}px", _baseRowHeight);
    }

    #endregion

    #region Public API

    /// <summary>
    /// Inicializuje row height manager s column information
    /// </summary>
    public void Initialize(List<GridColumnDefinition> columns, FontInfo fontInfo)
    {
        try
        {
            _fontInfo = fontInfo ?? new FontInfo();
            _columnWidths.Clear();

            // Setup column widths pre text measurement
            foreach (var column in columns)
            {
                var width = column.Width ?? column.MinWidth;
                _columnWidths[column.Name] = Math.Max(width, 50); // Min 50px width
            }

            _logger?.Info("📐 ROW HEIGHT MANAGER: Initialized with {Count} columns", columns.Count);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW HEIGHT ERROR: Initialization failed");
            throw;
        }
    }

    /// <summary>
    /// Recalculate unified row height pre dataset
    /// Všetky riadky dostanú výšku najvyššieho obsahu
    /// </summary>
    public async Task<double> RecalculateUnifiedRowHeightAsync(List<Dictionary<string, object?>> allRowData)
    {
        if (!_isEnabled) return _currentUnifiedRowHeight;

        try
        {
            _logger?.Info("📐 ROW HEIGHT CALC START: Processing {Count} rows for height calculation", allRowData.Count);

            double maxRequiredHeight = _baseRowHeight;

            // Phase 1: Measure content height pre každý riadok
            for (int rowIndex = 0; rowIndex < allRowData.Count; rowIndex++)
            {
                var rowData = allRowData[rowIndex];
                double rowRequiredHeight = await MeasureRowContentHeightAsync(rowData);

                if (rowRequiredHeight > maxRequiredHeight)
                {
                    maxRequiredHeight = rowRequiredHeight;
                    _logger?.Info("📐 ROW HEIGHT: New max height {Height}px found in row {Row}", 
                        Math.Ceiling(rowRequiredHeight), rowIndex);
                }
            }

            // Phase 2: Set unified height pre všetky riadky
            _currentUnifiedRowHeight = Math.Max(maxRequiredHeight, MinRowHeight);

            _logger?.Info("✅ ROW HEIGHT CALC SUCCESS: Unified height set to {Height}px for all rows", 
                Math.Ceiling(_currentUnifiedRowHeight));

            return _currentUnifiedRowHeight;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW HEIGHT ERROR: Height calculation failed");
            return _currentUnifiedRowHeight;
        }
    }

    /// <summary>
    /// Measure required height pre konkrétny cell content
    /// </summary>
    public async Task<double> MeasureCellContentHeightAsync(object? cellValue, string columnName)
    {
        if (!_isEnabled || cellValue == null) return _baseRowHeight;

        try
        {
            var cellText = cellValue.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(cellText)) return _baseRowHeight;

            var columnWidth = _columnWidths.TryGetValue(columnName, out double width) ? width : 120.0;
            
            var requiredHeight = await MeasureTextHeightAsync(cellText, columnWidth, _fontInfo);
            return Math.Max(requiredHeight, _baseRowHeight);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW HEIGHT ERROR: Cell content measurement failed - Column: {Column}", columnName);
            return _baseRowHeight;
        }
    }

    /// <summary>
    /// Quick recalculation pre nový content (single cell update)
    /// </summary>
    public async Task<bool> CheckIfRecalculationNeededAsync(object? newCellValue, string columnName)
    {
        if (!_isEnabled) return false;

        try
        {
            var requiredHeight = await MeasureCellContentHeightAsync(newCellValue, columnName);
            bool needsRecalculation = requiredHeight > _currentUnifiedRowHeight;

            if (needsRecalculation)
            {
                _logger?.Info("📐 ROW HEIGHT: Recalculation needed - new content requires {Height}px (current: {Current}px)", 
                    Math.Ceiling(requiredHeight), Math.Ceiling(_currentUnifiedRowHeight));
            }

            return needsRecalculation;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 ROW HEIGHT ERROR: Recalculation check failed");
            return false;
        }
    }

    /// <summary>
    /// Update column width (pre accurate text measurement)
    /// </summary>
    public void UpdateColumnWidth(string columnName, double newWidth)
    {
        _columnWidths[columnName] = Math.Max(newWidth, 50);
        _logger?.Info("📐 ROW HEIGHT: Column width updated - {Column}: {Width}px", columnName, newWidth);
    }

    /// <summary>
    /// Reset unified height to base height
    /// </summary>
    public void ResetToBaseHeight()
    {
        _currentUnifiedRowHeight = _baseRowHeight;
        _logger?.Info("📐 ROW HEIGHT: Reset to base height {Height}px", _baseRowHeight);
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Measure required height pre konkrétny riadok (všetky cells)
    /// </summary>
    private async Task<double> MeasureRowContentHeightAsync(Dictionary<string, object?> rowData)
    {
        double maxCellHeight = _baseRowHeight;

        foreach (var kvp in rowData)
        {
            var columnName = kvp.Key;
            var cellValue = kvp.Value;

            var cellHeight = await MeasureCellContentHeightAsync(cellValue, columnName);
            if (cellHeight > maxCellHeight)
            {
                maxCellHeight = cellHeight;
            }
        }

        return maxCellHeight;
    }

    /// <summary>
    /// Measure text height pre text content s danou šírkou
    /// </summary>
    private async Task<double> MeasureTextHeightAsync(string text, double availableWidth, FontInfo fontInfo)
    {
        try
        {
            // Create TextBlock pre measurement
            var textBlock = new TextBlock
            {
                Text = text,
                Width = availableWidth,
                TextWrapping = TextWrapping.Wrap,
                FontFamily = fontInfo.FontFamily,
                FontSize = fontInfo.FontSize,
                FontWeight = fontInfo.FontWeight,
                LineHeight = fontInfo.LineHeight
            };

            // Measure text size
            textBlock.Measure(new Size(availableWidth, double.PositiveInfinity));
            
            var measuredHeight = textBlock.DesiredSize.Height;
            
            // Add padding pre cell content (8px top + 8px bottom = 16px)
            var totalHeight = measuredHeight + 16;

            return Math.Max(totalHeight, _baseRowHeight);
        }
        catch (Exception)
        {
            // Fallback - ak measurement zlyhá
            var estimatedLines = Math.Ceiling(text.Length / (availableWidth / 8.0)); // Rough estimate
            var estimatedHeight = estimatedLines * 20 + 16; // 20px per line + padding
            return Math.Max(estimatedHeight, _baseRowHeight);
        }
    }

    #endregion
}

/// <summary>
/// Font information pre text measurement
/// </summary>
public class FontInfo
{
    /// <summary>
    /// Font family
    /// </summary>
    public Microsoft.UI.Xaml.Media.FontFamily FontFamily { get; set; } = new("Segoe UI");

    /// <summary>
    /// Font size
    /// </summary>
    public double FontSize { get; set; } = 14.0;

    /// <summary>
    /// Font weight
    /// </summary>
    public Windows.UI.Text.FontWeight FontWeight { get; set; } = FontWeights.Normal;

    /// <summary>
    /// Line height
    /// </summary>
    public double LineHeight { get; set; } = 20.0;
}