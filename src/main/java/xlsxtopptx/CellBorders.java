package xlsxtopptx;

/** The four border edges of a single cell. */
public record CellBorders(BorderEdgeStyle top, BorderEdgeStyle bottom, BorderEdgeStyle left, BorderEdgeStyle right) {
    public static final CellBorders NONE =
        new CellBorders(BorderEdgeStyle.NONE, BorderEdgeStyle.NONE, BorderEdgeStyle.NONE, BorderEdgeStyle.NONE);
}
