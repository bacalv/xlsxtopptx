package xlsxtopptx;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MergeIndexTest {

    // A wide merge (row 0, cols 3-4), a tall merge (rows 2-5, col 1), and one that's both
    // (rows 6-7, cols 2-3) -- covers every combination isCovered()/isWideAnchor()/isTallAnchor()
    // needs to distinguish.
    private final MergeIndex index = MergeIndex.build(List.of(
        new MergeRegion(0, 0, 3, 4),
        new MergeRegion(2, 5, 1, 1),
        new MergeRegion(6, 7, 2, 3)
    ));

    @Test
    void anchorCellOfAWideMergeIsNotItselfCovered() {
        assertFalse(index.isCovered(new CellPos(0, 3)));
        assertTrue(index.isWideAnchor(new CellPos(0, 3)));
    }

    @Test
    void trailingCellsOfAWideMergeAreCoveredButNotAnAnchor() {
        assertTrue(index.isCovered(new CellPos(0, 4)));
        assertFalse(index.isWideAnchor(new CellPos(0, 4)));
        assertFalse(index.isTallAnchor(new CellPos(0, 4)));
    }

    @Test
    void anchorCellOfATallMergeIsNotItselfCovered() {
        assertFalse(index.isCovered(new CellPos(2, 1)));
        assertTrue(index.isTallAnchor(new CellPos(2, 1)));
    }

    @Test
    void trailingCellsOfATallMergeAreCoveredButNotAnAnchor() {
        assertTrue(index.isCovered(new CellPos(5, 1)));
        assertFalse(index.isTallAnchor(new CellPos(5, 1)));
        assertFalse(index.isWideAnchor(new CellPos(5, 1)));
    }

    @Test
    void regionThatIsBothWideAndTallMarksBothAnchorSets() {
        var anchor = new CellPos(6, 2);

        assertTrue(index.isWideAnchor(anchor));
        assertTrue(index.isTallAnchor(anchor));
        assertFalse(index.isCovered(anchor));
        assertTrue(index.isCovered(new CellPos(7, 3)));
    }

    @Test
    void cellsOutsideAnyMergeAreUnaffected() {
        var untouched = new CellPos(20, 20);

        assertFalse(index.isCovered(untouched));
        assertFalse(index.isWideAnchor(untouched));
        assertFalse(index.isTallAnchor(untouched));
    }
}
