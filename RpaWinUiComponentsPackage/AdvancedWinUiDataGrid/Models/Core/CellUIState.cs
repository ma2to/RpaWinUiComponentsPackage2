using Windows.UI;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Core;

/// <summary>
/// UI stav bunky (farby, selection, validation, atď.) - part of HYBRID model
/// </summary>
public class CellUIState
{
    /// <summary>
    /// Je bunka označená pre kopírovanie
    /// </summary>
    public bool IsSelectedForCopy { get; set; }

    /// <summary>
    /// Je bunka momentálne editovaná
    /// </summary>
    public bool IsInEditMode { get; set; }

    /// <summary>
    /// Má bunka zmeny (dirty flag)
    /// </summary>
    public bool HasChanges { get; set; }

    /// <summary>
    /// Je bunka označená (selection)
    /// </summary>
    public bool IsSelected { get; set; }

    /// <summary>
    /// Má bunka focus
    /// </summary>
    public bool HasFocus { get; set; }

    /// <summary>
    /// Je bunka v hover stave
    /// </summary>
    public bool IsHovered { get; set; }

    /// <summary>
    /// Má bunka validation error
    /// </summary>
    public bool HasValidationError { get; set; }

    /// <summary>
    /// Validation error message (ak existuje)
    /// </summary>
    public string? ValidationErrorMessage { get; set; }

    /// <summary>
    /// Custom background color pre bunku
    /// </summary>
    public Color? CustomBackgroundColor { get; set; }

    /// <summary>
    /// Custom foreground color pre bunku
    /// </summary>
    public Color? CustomForegroundColor { get; set; }

    /// <summary>
    /// Custom border color pre bunku
    /// </summary>
    public Color? CustomBorderColor { get; set; }

    /// <summary>
    /// Reset všetkých UI stavov na default
    /// </summary>
    public void Reset()
    {
        IsSelectedForCopy = false;
        IsInEditMode = false;
        HasChanges = false;
        IsSelected = false;
        HasFocus = false;
        IsHovered = false;
        HasValidationError = false;
        ValidationErrorMessage = null;
        CustomBackgroundColor = null;
        CustomForegroundColor = null;
        CustomBorderColor = null;
    }

    /// <summary>
    /// Reset len selection stavu
    /// </summary>
    public void ResetSelection()
    {
        IsSelectedForCopy = false;
        IsSelected = false;
        HasFocus = false;
    }

    /// <summary>
    /// Reset len validation stavu
    /// </summary>
    public void ResetValidation()
    {
        HasValidationError = false;
        ValidationErrorMessage = null;
    }

    /// <summary>
    /// Reset len custom colors
    /// </summary>
    public void ResetColors()
    {
        CustomBackgroundColor = null;
        CustomForegroundColor = null;
        CustomBorderColor = null;
    }
}