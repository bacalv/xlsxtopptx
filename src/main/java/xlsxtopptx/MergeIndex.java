package xlsxtopptx;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Fast lookup over a sheet's merges: which cells are hidden inside a merge (covered), and which
 * anchor cells span multiple columns (wideAnchor) or multiple rows (tallAnchor) -- the latter two
 * matter because a merged cell's own column/row shouldn't be sized off its (visually stretched)
 * content the way an ordinary cell would.
 */
public record MergeIndex(Set<CellPos> covered, Set<CellPos> wideAnchor, Set<CellPos> tallAnchor) {

    public static MergeIndex build(List<MergeRegion> merges) {
        var covered = new HashSet<CellPos>();
        var wideAnchor = new HashSet<CellPos>();
        var tallAnchor = new HashSet<CellPos>();

        for (var merge : merges) {
            for (int r = merge.firstRow(); r <= merge.lastRow(); r++) {
                for (int c = merge.firstCol(); c <= merge.lastCol(); c++) {
                    if (r == merge.firstRow() && c == merge.firstCol()) continue;
                    covered.add(new CellPos(r, c));
                }
            }
            if (merge.isWide()) wideAnchor.add(merge.anchor());
            if (merge.isTall()) tallAnchor.add(merge.anchor());
        }

        return new MergeIndex(covered, wideAnchor, tallAnchor);
    }

    public boolean isCovered(CellPos pos) {
        return covered.contains(pos);
    }

    public boolean isWideAnchor(CellPos pos) {
        return wideAnchor.contains(pos);
    }

    public boolean isTallAnchor(CellPos pos) {
        return tallAnchor.contains(pos);
    }
}
