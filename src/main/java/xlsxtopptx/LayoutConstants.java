package xlsxtopptx;

import java.awt.Color;

/** Tunable constants governing the "scale everything to fit one slide" layout algorithm. */
public final class LayoutConstants {
    private LayoutConstants() {}

    public static final double MARGIN_PT = 24;
    public static final double TITLE_HEIGHT_PT = 26;
    // Bumped from 0.55 with extra headroom baked in: the true per-character width depends on the
    // font actually available in the viewer (Calibri, a substituted sans, etc.), which we can't
    // know in advance, so this trades a little unused column width for insurance against wrapping
    // in the widest cell of a column.
    public static final double CHAR_WIDTH_FACTOR = 0.62;
    public static final double CELL_PADDING_PT = 8;
    public static final double LINE_HEIGHT_FACTOR = 1.25;
    public static final double ROW_PADDING_PT = 4;
    public static final double MIN_FONT_PT = 4.0;
    public static final double DEFAULT_FONT_PT = 11.0;
    public static final double MIN_COL_WIDTH_PT = 14;
    public static final double MIN_ROW_HEIGHT_PT = 10;
    // Reserved beyond the geometric margins: our width/height estimates are necessarily
    // approximate (real text metrics depend on the font actually available in whatever app opens
    // the file), so a small cushion keeps a slightly-off estimate from pushing the last row or
    // column past the visible slide edge.
    public static final double OVERFLOW_SAFETY_PT = 12;
    public static final Color DEFAULT_BORDER_COLOR = Color.BLACK;
    public static final Color DEFAULT_FONT_COLOR = Color.BLACK;
    public static final String DEFAULT_FONT_FAMILY = "Calibri";

    public static final double LINK_HEIGHT_PT = 16;
    public static final double LINK_FONT_PT = 9.0;
    public static final Color LINK_FONT_COLOR = new Color(0x05, 0x63, 0xC1);
}
