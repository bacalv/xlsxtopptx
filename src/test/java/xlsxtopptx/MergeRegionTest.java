package xlsxtopptx;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MergeRegionTest {

    @Test
    void singleCellRegionIsNeitherWideNorTall() {
        var region = new MergeRegion(2, 2, 3, 3);

        assertFalse(region.isWide());
        assertFalse(region.isTall());
        assertFalse(region.spansMultipleCells());
    }

    @Test
    void wideRegionSpansMultipleColumnsInOneRow() {
        var region = new MergeRegion(0, 0, 3, 4);

        assertTrue(region.isWide());
        assertFalse(region.isTall());
        assertTrue(region.spansMultipleCells());
    }

    @Test
    void tallRegionSpansMultipleRowsInOneColumn() {
        var region = new MergeRegion(1, 5, 2, 2);

        assertFalse(region.isWide());
        assertTrue(region.isTall());
        assertTrue(region.spansMultipleCells());
    }

    @Test
    void anchorIsTheTopLeftCell() {
        var region = new MergeRegion(4, 6, 7, 9);

        assertEquals(new CellPos(4, 7), region.anchor());
    }
}
