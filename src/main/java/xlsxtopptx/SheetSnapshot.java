package xlsxtopptx;

import java.util.List;

/** The full, POI-independent snapshot of one Excel sheet: its rows, merges, and column count. */
public record SheetSnapshot(List<RowSnapshot> rows, List<MergeRegion> merges, int numCols) {

    public int rowCount() {
        return rows.size();
    }
}
