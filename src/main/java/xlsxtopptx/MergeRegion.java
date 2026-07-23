package xlsxtopptx;

/** A merged (or "center across selection") cell range, in zero-based rows/columns, inclusive. */
public record MergeRegion(int firstRow, int lastRow, int firstCol, int lastCol) {

    public boolean isWide() {
        return lastCol > firstCol;
    }

    public boolean isTall() {
        return lastRow > firstRow;
    }

    public boolean spansMultipleCells() {
        return isWide() || isTall();
    }

    /** The top-left cell of the region -- the only one that actually renders content. */
    public CellPos anchor() {
        return new CellPos(firstRow, firstCol);
    }
}
