package xlsxtopptx;

import org.apache.poi.ss.usermodel.BorderStyle;

import java.awt.Color;

/** The style and color of a single border edge (top, bottom, left, or right) of a cell. */
public record BorderEdgeStyle(BorderStyle style, Color color) {
    public static final BorderEdgeStyle NONE = new BorderEdgeStyle(BorderStyle.NONE, null);

    public boolean isVisible() {
        return style != null && style != BorderStyle.NONE;
    }
}
