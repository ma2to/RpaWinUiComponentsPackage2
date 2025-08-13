namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;

/// <summary>
/// Pozícia bunky v gridu (row, column)
/// </summary>
public struct CellPosition : IEquatable<CellPosition>
{
    /// <summary>
    /// Riadok (0-based index)
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// Stĺpec (0-based index)
    /// </summary>
    public int Column { get; set; }

    public CellPosition(int row, int column)
    {
        Row = row;
        Column = column;
    }

    /// <summary>
    /// Neplatná pozícia (-1, -1)
    /// </summary>
    public static CellPosition Invalid => new(-1, -1);

    /// <summary>
    /// Nulová pozícia (0, 0)
    /// </summary>
    public static CellPosition Zero => new(0, 0);

    /// <summary>
    /// Je pozícia platná (>= 0)
    /// </summary>
    public bool IsValid => Row >= 0 && Column >= 0;

    public bool Equals(CellPosition other)
    {
        return Row == other.Row && Column == other.Column;
    }

    public override bool Equals(object? obj)
    {
        return obj is CellPosition other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Row, Column);
    }

    public static bool operator ==(CellPosition left, CellPosition right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(CellPosition left, CellPosition right)
    {
        return !left.Equals(right);
    }

    public override string ToString()
    {
        return $"({Row}, {Column})";
    }
}