package xlsxtopptx;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.junit.jupiter.api.Test;

import java.awt.Color;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class BorderEdgeStyleTest {

    @Test
    void noneConstantIsNotVisible() {
        assertFalse(BorderEdgeStyle.NONE.isVisible());
    }

    @Test
    void explicitNoneStyleIsNotVisible() {
        assertFalse(new BorderEdgeStyle(BorderStyle.NONE, Color.BLACK).isVisible());
    }

    @Test
    void nullStyleIsNotVisible() {
        // Defensive: a cell that was never given an explicit border style at all.
        assertFalse(new BorderEdgeStyle(null, null).isVisible());
    }

    @Test
    void aRealBorderStyleIsVisible() {
        assertTrue(new BorderEdgeStyle(BorderStyle.THIN, Color.BLUE).isVisible());
    }
}
