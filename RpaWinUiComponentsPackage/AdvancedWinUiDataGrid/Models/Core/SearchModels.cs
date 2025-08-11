using System.Text.RegularExpressions;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Core;

/// <summary>
/// Výsledok search operácie
/// </summary>
public class SearchResults
{
    public List<SearchMatch> Matches { get; set; } = new();
    public int TotalMatchCount => Matches.Count;
    public string SearchTerm { get; set; } = string.Empty;
    public SearchConfiguration Configuration { get; set; } = new();
    public TimeSpan SearchDuration { get; set; }
    public bool HasMatches => Matches.Count > 0;
}

/// <summary>
/// Konkrétny match v search operácii
/// </summary>
public class SearchMatch
{
    public int RowIndex { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public object? MatchedValue { get; set; }
    public string MatchedText { get; set; } = string.Empty;
    public int MatchStartIndex { get; set; }
    public int MatchLength { get; set; }
}

/// <summary>
/// Konfigurácia pre search operáciu
/// </summary>
public class SearchConfiguration
{
    /// <summary>
    /// Target columns pre search (null = všetky columns)
    /// </summary>
    public List<string>? TargetColumns { get; set; }

    /// <summary>
    /// Target rows pre search (null = všetky rows)
    /// </summary>
    public List<int>? TargetRows { get; set; }

    /// <summary>
    /// Case sensitive search
    /// </summary>
    public bool CaseSensitive { get; set; } = false;

    /// <summary>
    /// Regular expression search
    /// </summary>
    public bool IsRegex { get; set; } = false;

    /// <summary>
    /// Whole word search
    /// </summary>
    public bool WholeWord { get; set; } = false;

    /// <summary>
    /// Search mode
    /// </summary>
    public SearchMode Mode { get; set; } = SearchMode.Contains;
}

/// <summary>
/// Search mode options
/// </summary>
public enum SearchMode
{
    Contains,
    StartsWith,
    EndsWith,
    Exact
}

