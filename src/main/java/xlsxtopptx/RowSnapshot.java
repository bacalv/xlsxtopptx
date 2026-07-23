package xlsxtopptx;

import java.util.List;

/** One row's worth of cell snapshots, always padded out to the sheet's full column count. */
public record RowSnapshot(List<CellSnapshot> cells) {

    public CellSnapshot cellAt(int col) {
        return col < cells.size() ? cells.get(col) : CellSnapshot.empty();
    }
}
