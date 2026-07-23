package xlsxtopptx;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;

import java.awt.Color;

/** An immutable, POI-independent snapshot of everything one Excel cell needs rendered. */
public record CellSnapshot(
    String text,
    Color fill,
    Color fontColor,
    boolean bold,
    boolean italic,
    double fontSize,
    String fontName,
    TextAlign align,
    CellBorders borders
) {
    public static CellSnapshot empty() {
        return new CellSnapshot(
            "", null, LayoutConstants.DEFAULT_FONT_COLOR, false, false,
            LayoutConstants.DEFAULT_FONT_PT, null, TextAlign.LEFT, CellBorders.NONE);
    }

    public boolean isEmpty() {
        return text == null || text.isEmpty();
    }
}
