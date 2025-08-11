using Microsoft.Extensions.Logging.Abstractions;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Models.Configuration;
using RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Utilities.Helpers;

namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Services.Core;

/// <summary>
/// Smart resolver pre duplicate column names s Reserved Name Protection
/// Special columns majú vždy prioritu pri naming, user columns sa automaticky premenujú
/// Business logika používa flag-based lookup (IsSpecialColumn), nie name-based
/// </summary>
public class SmartColumnNameResolver
{
    #region Private Fields

    /// <summary>
    /// Reserved names pre special columns (priority names)
    /// </summary>
    private readonly HashSet<string> _reservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "ValidationAlerts", 
        "DeleteRow", 
        "DeleteRows",
        "CheckBox"
    };

    /// <summary>
    /// Logger (nullable - funguje bez loggera)
    /// </summary>
    private readonly Microsoft.Extensions.Logging.ILogger? _logger;

    #endregion

    #region Constructor

    /// <summary>
    /// Konštruktor s optional logger
    /// </summary>
    public SmartColumnNameResolver(Microsoft.Extensions.Logging.ILogger? logger = null)
    {
        _logger = logger;
    }

    #endregion

    #region Public API

    /// <summary>
    /// Resolve duplicate column names s Reserved Name Protection
    /// Special columns dostanú svoje preferované názvy, user columns sa premenujú
    /// </summary>
    /// <param name="columns">Pôvodné column definitions s možnými duplicitmi</param>
    /// <returns>Resolved column definitions bez duplicitov</returns>
    public List<GridColumnDefinition> ResolveDuplicateNames(List<GridColumnDefinition> columns)
    {
        try
        {
            _logger?.Info("🔧 COLUMN RESOLVER START: Processing {Count} columns for duplicate resolution", columns.Count);

            var resolvedColumns = new List<GridColumnDefinition>();
            var usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Phase 1: Process special columns first (priority naming)
            var specialColumns = columns.Where(c => c.IsSpecialColumn).ToList();
            var userColumns = columns.Where(c => !c.IsSpecialColumn).ToList();

            _logger?.Info("🔧 COLUMN RESOLVER: Special columns: {SpecialCount}, User columns: {UserCount}", 
                specialColumns.Count, userColumns.Count);

            // Process special columns - oni majú prioritu na svoje reserved names
            foreach (var specialColumn in specialColumns)
            {
                var resolvedColumn = specialColumn.Clone();
                var resolvedName = ResolveSpecialColumnName(resolvedColumn, usedNames);
                
                resolvedColumn.Name = resolvedName;
                resolvedColumns.Add(resolvedColumn);

                _logger?.Info("🔧 SPECIAL COLUMN: '{OriginalName}' → '{ResolvedName}' ({Type})", 
                    specialColumn.Name, resolvedName, GetSpecialColumnType(specialColumn));
            }

            // Phase 2: Process user columns (môžu sa premenúvať)
            foreach (var userColumn in userColumns)
            {
                var resolvedColumn = userColumn.Clone();
                var resolvedName = ResolveUserColumnName(resolvedColumn, usedNames);
                
                resolvedColumn.Name = resolvedName;
                resolvedColumns.Add(resolvedColumn);

                if (resolvedName != userColumn.Name)
                {
                    _logger?.Info("🔧 USER COLUMN RENAMED: '{OriginalName}' → '{ResolvedName}' (duplicate resolved)", 
                        userColumn.Name, resolvedName);
                }
                else
                {
                    _logger?.Info("🔧 USER COLUMN: '{Name}' (no change needed)", userColumn.Name);
                }
            }

            // Validation check - žiadne duplicates nesmú zostať
            var finalNames = resolvedColumns.Select(c => c.Name).ToList();
            var duplicatesCheck = finalNames.GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                                          .Where(g => g.Count() > 1)
                                          .Select(g => g.Key)
                                          .ToList();

            if (duplicatesCheck.Any())
            {
                _logger?.Error("🚨 COLUMN RESOLVER ERROR: Unresolved duplicates found: {Duplicates}", 
                    string.Join(", ", duplicatesCheck));
                throw new InvalidOperationException($"Column name resolution failed - unresolved duplicates: {string.Join(", ", duplicatesCheck)}");
            }

            _logger?.Info("✅ COLUMN RESOLVER SUCCESS: {Count} columns resolved, no duplicates", resolvedColumns.Count);
            return resolvedColumns;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "🚨 COLUMN RESOLVER ERROR: Duplicate resolution failed");
            throw;
        }
    }

    /// <summary>
    /// Získa finálny zoznam názvov columns po resolve (pre API)
    /// </summary>
    public List<string> GetResolvedColumnNames(List<GridColumnDefinition> originalColumns)
    {
        var resolved = ResolveDuplicateNames(originalColumns);
        return resolved.Select(c => c.Name).ToList();
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Resolve názov pre special column (priority naming)
    /// </summary>
    private string ResolveSpecialColumnName(GridColumnDefinition specialColumn, Dictionary<string, int> usedNames)
    {
        var preferredName = GetPreferredSpecialColumnName(specialColumn);
        
        // Special columns si môžu rezervovať svoje preferované názvy
        if (!usedNames.ContainsKey(preferredName))
        {
            usedNames[preferredName] = 1;
            return preferredName;
        }

        // Ak je preferovaný názov už použitý, použije sa numbering
        var resolvedName = GenerateUniqueColumnName(preferredName, usedNames);
        return resolvedName;
    }

    /// <summary>
    /// Resolve názov pre user column (môže sa premenúvať)
    /// </summary>
    private string ResolveUserColumnName(GridColumnDefinition userColumn, Dictionary<string, int> usedNames)
    {
        var originalName = userColumn.Name;

        // Ak názov nie je použitý, jednoducho ho použijem
        if (!usedNames.ContainsKey(originalName))
        {
            usedNames[originalName] = 1;
            return originalName;
        }

        // Ak je názov už použitý, wygeneruje sa unique variant
        var resolvedName = GenerateUniqueColumnName(originalName, usedNames);
        return resolvedName;
    }

    /// <summary>
    /// Generuje unique column name s numerickými suffixmi
    /// </summary>
    private string GenerateUniqueColumnName(string baseName, Dictionary<string, int> usedNames)
    {
        string uniqueName;
        int counter = 2; // Začína s "_2" (pôvodný je považovaný za "_1")

        do
        {
            uniqueName = $"{baseName}_{counter}";
            counter++;
        }
        while (usedNames.ContainsKey(uniqueName));

        usedNames[uniqueName] = 1;
        return uniqueName;
    }

    /// <summary>
    /// Získa preferovaný názov pre special column
    /// </summary>
    private string GetPreferredSpecialColumnName(GridColumnDefinition specialColumn)
    {
        if (specialColumn.IsValidationAlertsColumn)
            return "ValidationAlerts";
        
        if (specialColumn.IsDeleteRowColumn)
            return "DeleteRow";
        
        if (specialColumn.IsCheckBoxColumn)
            return "CheckBox";

        // Fallback - ak nie je rozoznané ako special column
        return specialColumn.Name;
    }

    /// <summary>
    /// Získa typ special column pre logging
    /// </summary>
    private string GetSpecialColumnType(GridColumnDefinition specialColumn)
    {
        if (specialColumn.IsValidationAlertsColumn) return "ValidationAlerts";
        if (specialColumn.IsDeleteRowColumn) return "DeleteRow";
        if (specialColumn.IsCheckBoxColumn) return "CheckBox";
        return "Unknown";
    }

    #endregion

    #region Validation Helpers

    /// <summary>
    /// Validuje resolved columns (pre unit testing)
    /// </summary>
    public bool ValidateResolvedColumns(List<GridColumnDefinition> resolvedColumns, out List<string> errors)
    {
        errors = new List<string>();

        // Check for duplicates
        var names = resolvedColumns.Select(c => c.Name).ToList();
        var duplicates = names.GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                             .Where(g => g.Count() > 1)
                             .Select(g => g.Key)
                             .ToList();

        if (duplicates.Any())
        {
            errors.Add($"Duplicate names found: {string.Join(", ", duplicates)}");
        }

        // Check special columns have proper naming
        foreach (var column in resolvedColumns.Where(c => c.IsSpecialColumn))
        {
            var preferredName = GetPreferredSpecialColumnName(column);
            var namePattern = $"^{preferredName}(_\\d+)?$";
            
            if (!System.Text.RegularExpressions.Regex.IsMatch(column.Name, namePattern))
            {
                errors.Add($"Special column '{column.Name}' doesn't follow proper naming pattern '{preferredName}' or '{preferredName}_X'");
            }
        }

        return errors.Count == 0;
    }

    #endregion
}