namespace RpaWinUiComponentsPackage.AdvancedWinUiDataGrid.Modules.Table.Models;

/// <summary>
/// Rozsah buniek v gridu (od-do pozície)
/// </summary>
public struct CellRange : IEquatable<CellRange>
{
    /// <summary>
    /// Začiatočná pozícia (ľavý horný roh)
    /// </summary>
    public CellPosition Start { get; set; }

    /// <summary>
    /// Koncová pozícia (pravý dolný roh)
    /// </summary>
    public CellPosition End { get; set; }

    public CellRange(CellPosition start, CellPosition end)
    {
        // Zabezpečíme správne poradie (Start <= End)
        Start = new CellPosition(
            Math.Min(start.Row, end.Row),
            Math.Min(start.Column, end.Column)
        );
        End = new CellPosition(
            Math.Max(start.Row, end.Row),
            Math.Max(start.Column, end.Column)
        );
    }

    public CellRange(int startRow, int startColumn, int endRow, int endColumn)
        : this(new CellPosition(startRow, startColumn), new CellPosition(endRow, endColumn))
    {
    }

    /// <summary>
    /// Prázdny range
    /// </summary>
    public static CellRange Empty => new(CellPosition.Invalid, CellPosition.Invalid);

    /// <summary>
    /// Single cell range
    /// </summary>
    public static CellRange Single(CellPosition position) => new(position, position);

    /// <summary>
    /// Single cell range
    /// </summary>
    public static CellRange Single(int row, int column) => new(row, column, row, column);

    /// <summary>
    /// Je range prázdny
    /// </summary>
    public bool IsEmpty => !Start.IsValid || !End.IsValid;

    /// <summary>
    /// Je range single cell
    /// </summary>
    public bool IsSingleCell => Start == End;

    /// <summary>
    /// Počet riadkov v range
    /// </summary>
    public int RowCount => IsEmpty ? 0 : End.Row - Start.Row + 1;

    /// <summary>
    /// Počet stĺpcov v range
    /// </summary>
    public int ColumnCount => IsEmpty ? 0 : End.Column - Start.Column + 1;

    /// <summary>
    /// Celkový počet buniek v range
    /// </summary>
    public int TotalCells => RowCount * ColumnCount;

    /// <summary>
    /// Obsahuje range danú pozíciu
    /// </summary>
    public bool Contains(CellPosition position)
    {
        return !IsEmpty && 
               position.Row >= Start.Row && position.Row <= End.Row &&
               position.Column >= Start.Column && position.Column <= End.Column;
    }

    /// <summary>
    /// Obsahuje range danú pozíciu
    /// </summary>
    public bool Contains(int row, int column)
    {
        return Contains(new CellPosition(row, column));
    }

    /// <summary>
    /// Prekrývajú sa dva ranges
    /// </summary>
    public bool Intersects(CellRange other)
    {
        return !IsEmpty && !other.IsEmpty &&
               Start.Row <= other.End.Row && End.Row >= other.Start.Row &&
               Start.Column <= other.End.Column && End.Column >= other.Start.Column;
    }

    /// <summary>
    /// Priesečník dvoch ranges
    /// </summary>
    public CellRange Intersection(CellRange other)
    {
        if (!Intersects(other))
            return Empty;

        return new CellRange(
            new CellPosition(Math.Max(Start.Row, other.Start.Row), Math.Max(Start.Column, other.Start.Column)),
            new CellPosition(Math.Min(End.Row, other.End.Row), Math.Min(End.Column, other.Start.Column))
        );
    }

    /// <summary>
    /// Rozšírenie range o danú pozíciu
    /// </summary>
    public CellRange Expand(CellPosition position)
    {
        if (IsEmpty)
            return Single(position);

        return new CellRange(
            new CellPosition(Math.Min(Start.Row, position.Row), Math.Min(Start.Column, position.Column)),
            new CellPosition(Math.Max(End.Row, position.Row), Math.Max(End.Column, position.Column))
        );
    }

    /// <summary>
    /// Enumerate všetkých pozícií v range
    /// </summary>
    public IEnumerable<CellPosition> GetAllPositions()
    {
        if (IsEmpty)
            yield break;

        for (int row = Start.Row; row <= End.Row; row++)
        {
            for (int col = Start.Column; col <= End.Column; col++)
            {
                yield return new CellPosition(row, col);
            }
        }
    }

    public bool Equals(CellRange other)
    {
        return Start.Equals(other.Start) && End.Equals(other.End);
    }

    public override bool Equals(object? obj)
    {
        return obj is CellRange other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Start, End);
    }

    public static bool operator ==(CellRange left, CellRange right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(CellRange left, CellRange right)
    {
        return !(left == right);
    }

    public override string ToString()
    {
        return IsEmpty ? "Empty" : $"{Start} -> {End}";
    }
}