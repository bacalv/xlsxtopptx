package xlsxtopptx;

import java.util.List;

/**
 * The result of {@link LayoutEngine#compute}: the uniform font scale applied to every cell, the
 * final column widths and row heights (in points), and the resulting overall table size.
 */
public record SheetLayout(
    double fontScale,
    List<Double> colWidths,
    List<Double> rowHeights,
    double tableWidth,
    double tableHeight
) {
}
