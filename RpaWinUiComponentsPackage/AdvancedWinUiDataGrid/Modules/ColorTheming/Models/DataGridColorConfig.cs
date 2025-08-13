using Windows.UI;
using Microsoft.UI;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.ColorTheming.Models;

/// <summary>
/// Konfigurácia farieb pre DataGrid - runtime color theming
/// </summary>
public class DataGridColorConfig
{
    /// <summary>
    /// Farba pozadia buniek (default)
    /// </summary>
    public Color? CellBackgroundColor { get; set; }

    /// <summary>
    /// Farba textu v bunkách (default)
    /// </summary>
    public Color? CellForegroundColor { get; set; }

    /// <summary>
    /// Farba orámovani buniek
    /// </summary>
    public Color? CellBorderColor { get; set; }

    /// <summary>
    /// Farba pozadia pre označené bunky
    /// </summary>
    public Color? SelectionBackgroundColor { get; set; }

    /// <summary>
    /// Farba textu pre označené bunky
    /// </summary>
    public Color? SelectionForegroundColor { get; set; }

    /// <summary>
    /// Farba pozadia pre bunky v copy mode
    /// </summary>
    public Color? CopyModeBackgroundColor { get; set; }

    /// <summary>
    /// Farba orámovani pre validation errors
    /// </summary>
    public Color? ValidationErrorBorderColor { get; set; }

    /// <summary>
    /// Farba pozadia pre validation errors
    /// </summary>
    public Color? ValidationErrorBackgroundColor { get; set; }

    /// <summary>
    /// Farba pre hover efekt
    /// </summary>
    public Color? HoverBackgroundColor { get; set; }

    /// <summary>
    /// Farba pre focus ring
    /// </summary>
    public Color? FocusRingColor { get; set; }

    /// <summary>
    /// Farba pozadia pre párne riadky (zebra pattern)
    /// </summary>
    public Color? EvenRowBackgroundColor { get; set; }

    /// <summary>
    /// Farba pozadia pre nepárne riadky (zebra pattern)
    /// </summary>
    public Color? OddRowBackgroundColor { get; set; }

    /// <summary>
    /// Farba pozadia pre header
    /// </summary>
    public Color? HeaderBackgroundColor { get; set; }

    /// <summary>
    /// Farba textu pre header
    /// </summary>
    public Color? HeaderForegroundColor { get; set; }

    /// <summary>
    /// Farba orámovani pre header
    /// </summary>
    public Color? HeaderBorderColor { get; set; }

    /// <summary>
    /// Default color scheme (Windows 11 style) - UPDATED per user requirements
    /// Focus: bledo zelený, Copy: bledo modrý, Background: biely/zebra, Border: čierny, Validation error: červený
    /// </summary>
    public static DataGridColorConfig Default => new()
    {
        CellBackgroundColor = Colors.White,
        CellForegroundColor = Colors.Black,
        CellBorderColor = Colors.Black, // UPDATED: čierny border namiesto šedý
        SelectionBackgroundColor = Color.FromArgb(100, 144, 238, 144), // UPDATED: bledo zelený focus namiesto modrý
        SelectionForegroundColor = Colors.Black,
        CopyModeBackgroundColor = Color.FromArgb(100, 173, 216, 230), // UPDATED: bledo modrý copy mode
        ValidationErrorBorderColor = Colors.Red,
        ValidationErrorBackgroundColor = Color.FromArgb(50, 255, 0, 0), // Light red with transparency
        HoverBackgroundColor = Color.FromArgb(50, 0, 0, 0), // Light gray with transparency
        FocusRingColor = Color.FromArgb(255, 144, 238, 144), // UPDATED: bledo zelený focus ring
        EvenRowBackgroundColor = Colors.White,
        OddRowBackgroundColor = Color.FromArgb(255, 249, 249, 249), // Very light gray - zebra pattern
        HeaderBackgroundColor = Color.FromArgb(255, 240, 240, 240), // Light gray
        HeaderForegroundColor = Colors.Black,
        HeaderBorderColor = Colors.Black // UPDATED: čierny header border namiesto šedý
    };

    /// <summary>
    /// Dark theme color scheme
    /// </summary>
    public static DataGridColorConfig Dark => new()
    {
        CellBackgroundColor = Color.FromArgb(255, 32, 32, 32), // Dark gray
        CellForegroundColor = Colors.White,
        CellBorderColor = Color.FromArgb(255, 60, 60, 60), // Medium gray
        SelectionBackgroundColor = Color.FromArgb(100, 0, 120, 215), // Light blue with transparency
        SelectionForegroundColor = Colors.White,
        CopyModeBackgroundColor = Color.FromArgb(100, 70, 130, 180), // Steel blue with transparency
        ValidationErrorBorderColor = Colors.Red,
        ValidationErrorBackgroundColor = Color.FromArgb(50, 255, 0, 0), // Light red with transparency
        HoverBackgroundColor = Color.FromArgb(50, 255, 255, 255), // Light white with transparency
        FocusRingColor = Color.FromArgb(255, 0, 120, 215), // Windows accent blue
        EvenRowBackgroundColor = Color.FromArgb(255, 32, 32, 32), // Dark gray
        OddRowBackgroundColor = Color.FromArgb(255, 28, 28, 28), // Slightly darker gray
        HeaderBackgroundColor = Color.FromArgb(255, 40, 40, 40), // Medium dark gray
        HeaderForegroundColor = Colors.White,
        HeaderBorderColor = Color.FromArgb(255, 60, 60, 60)
    };

    /// <summary>
    /// Klonuje color config
    /// </summary>
    public DataGridColorConfig Clone()
    {
        return new DataGridColorConfig
        {
            CellBackgroundColor = CellBackgroundColor,
            CellForegroundColor = CellForegroundColor,
            CellBorderColor = CellBorderColor,
            SelectionBackgroundColor = SelectionBackgroundColor,
            SelectionForegroundColor = SelectionForegroundColor,
            CopyModeBackgroundColor = CopyModeBackgroundColor,
            ValidationErrorBorderColor = ValidationErrorBorderColor,
            ValidationErrorBackgroundColor = ValidationErrorBackgroundColor,
            HoverBackgroundColor = HoverBackgroundColor,
            FocusRingColor = FocusRingColor,
            EvenRowBackgroundColor = EvenRowBackgroundColor,
            OddRowBackgroundColor = OddRowBackgroundColor,
            HeaderBackgroundColor = HeaderBackgroundColor,
            HeaderForegroundColor = HeaderForegroundColor,
            HeaderBorderColor = HeaderBorderColor
        };
    }

    /// <summary>
    /// Merge s iným color config (non-null values override)
    /// </summary>
    public void MergeWith(DataGridColorConfig other)
    {
        if (other.CellBackgroundColor.HasValue) CellBackgroundColor = other.CellBackgroundColor;
        if (other.CellForegroundColor.HasValue) CellForegroundColor = other.CellForegroundColor;
        if (other.CellBorderColor.HasValue) CellBorderColor = other.CellBorderColor;
        if (other.SelectionBackgroundColor.HasValue) SelectionBackgroundColor = other.SelectionBackgroundColor;
        if (other.SelectionForegroundColor.HasValue) SelectionForegroundColor = other.SelectionForegroundColor;
        if (other.CopyModeBackgroundColor.HasValue) CopyModeBackgroundColor = other.CopyModeBackgroundColor;
        if (other.ValidationErrorBorderColor.HasValue) ValidationErrorBorderColor = other.ValidationErrorBorderColor;
        if (other.ValidationErrorBackgroundColor.HasValue) ValidationErrorBackgroundColor = other.ValidationErrorBackgroundColor;
        if (other.HoverBackgroundColor.HasValue) HoverBackgroundColor = other.HoverBackgroundColor;
        if (other.FocusRingColor.HasValue) FocusRingColor = other.FocusRingColor;
        if (other.EvenRowBackgroundColor.HasValue) EvenRowBackgroundColor = other.EvenRowBackgroundColor;
        if (other.OddRowBackgroundColor.HasValue) OddRowBackgroundColor = other.OddRowBackgroundColor;
        if (other.HeaderBackgroundColor.HasValue) HeaderBackgroundColor = other.HeaderBackgroundColor;
        if (other.HeaderForegroundColor.HasValue) HeaderForegroundColor = other.HeaderForegroundColor;
        if (other.HeaderBorderColor.HasValue) HeaderBorderColor = other.HeaderBorderColor;
    }
}