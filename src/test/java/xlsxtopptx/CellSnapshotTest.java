package xlsxtopptx;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class CellSnapshotTest {

    @Test
    void emptyFactoryProducesABlankCell() {
        var cell = CellSnapshot.empty();

        assertTrue(cell.isEmpty());
        assertEquals("", cell.text());
        assertEquals(LayoutConstants.DEFAULT_FONT_PT, cell.fontSize());
        assertEquals(TextAlign.LEFT, cell.align());
        assertEquals(CellBorders.NONE, cell.borders());
    }

    @Test
    void cellWithTextIsNotEmpty() {
        var cell = new CellSnapshot("Hello", null, LayoutConstants.DEFAULT_FONT_COLOR, false, false,
            11.0, "Calibri", TextAlign.LEFT, CellBorders.NONE);

        assertFalse(cell.isEmpty());
    }

    @Test
    void cellWithNullTextIsTreatedAsEmpty() {
        var cell = new CellSnapshot(null, null, LayoutConstants.DEFAULT_FONT_COLOR, false, false,
            11.0, "Calibri", TextAlign.LEFT, CellBorders.NONE);

        assertTrue(cell.isEmpty());
    }

    @Test
    void cellWithOnlyWhitespaceIsNotConsideredEmpty() {
        // isEmpty() checks for a zero-length string, not blank -- a single space is real content
        // as far as layout is concerned.
        var cell = new CellSnapshot(" ", null, LayoutConstants.DEFAULT_FONT_COLOR, false, false,
            11.0, "Calibri", TextAlign.LEFT, CellBorders.NONE);

        assertFalse(cell.isEmpty());
    }
}
