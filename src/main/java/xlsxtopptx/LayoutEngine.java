package xlsxtopptx;

import java.util.ArrayList;
import java.util.List;

/**
 * Computes a single uniform font scale (and the column widths / row heights that follow from it)
 * so an entire sheet's content fits within the available slide area -- content-and-font-aware,
 * scaled down only as far as needed to fit, and never stretched back up beyond a sheet's natural
 * size just to fill leftover space.
 */
public final class LayoutEngine {
    private LayoutEngine() {}

    /** Equivalent to {@link #compute(SheetSnapshot, MergeIndex, double, double, boolean, boolean)}
     *  with both scale-to-fit flags off -- the sheet keeps its natural width/height instead of
     *  being stretched to fill the available space. */
    public static SheetLayout compute(SheetSnapshot sheet, MergeIndex mergeIndex, double availableWidth, double availableHeight) {
        return compute(sheet, mergeIndex, availableWidth, availableHeight, false, false);
    }

    public static SheetLayout compute(SheetSnapshot sheet, MergeIndex mergeIndex, double availableWidth, double availableHeight,
                                       boolean scaleToFitWidth, boolean scaleToFitHeight) {
        var rows = sheet.rows();
        var numCols = sheet.numCols();

        var totalW1x = 0.0;
        for (int c = 0; c < numCols; c++) totalW1x += naturalColWidth(rows, c, 1.0, mergeIndex);
        var totalH1x = 0.0;
        for (int r = 0; r < rows.size(); r++) totalH1x += naturalRowHeight(rows.get(r), r, 1.0, mergeIndex);

        var minFontInSheet = LayoutConstants.DEFAULT_FONT_PT;
        for (var row : rows) {
            for (var cell : row.cells()) {
                if (!cell.isEmpty()) {
                    minFontInSheet = Math.min(minFontInSheet, cell.fontSize() > 0 ? cell.fontSize() : LayoutConstants.DEFAULT_FONT_PT);
                }
            }
        }

        var scaleW = totalW1x > 0 ? availableWidth / totalW1x : 1.0;
        var scaleH = totalH1x > 0 ? availableHeight / totalH1x : 1.0;
        var fontScale = Math.min(1.0, Math.min(scaleW, scaleH));
        var minScaleFloor = LayoutConstants.MIN_FONT_PT / minFontInSheet;
        fontScale = Math.max(fontScale, minScaleFloor);

        var colWidths = new double[numCols];
        var wSum = 0.0;
        for (int c = 0; c < numCols; c++) {
            colWidths[c] = naturalColWidth(rows, c, fontScale, mergeIndex);
            wSum += colWidths[c];
        }
        // Per-column/row minimum-size floors mean the natural size can still exceed the available
        // space even after fontScale was chosen (e.g. many short rows each clamped at
        // MIN_ROW_HEIGHT_PT) -- so a shrink-if-needed safety net always applies. Growing past the
        // natural size to fill the available space is opt-in via scaleToFitWidth/scaleToFitHeight;
        // left off, the table just keeps its natural (possibly smaller) size.
        var wGrow = wSum > 0 ? (scaleToFitWidth ? availableWidth / wSum : Math.min(1.0, availableWidth / wSum)) : 1.0;
        for (int c = 0; c < numCols; c++) colWidths[c] *= wGrow;

        var rowHeights = new double[rows.size()];
        var hSum = 0.0;
        for (int r = 0; r < rows.size(); r++) {
            rowHeights[r] = naturalRowHeight(rows.get(r), r, fontScale, mergeIndex);
            hSum += rowHeights[r];
        }
        var hGrow = hSum > 0 ? (scaleToFitHeight ? availableHeight / hSum : Math.min(1.0, availableHeight / hSum)) : 1.0;
        for (int r = 0; r < rows.size(); r++) rowHeights[r] *= hGrow;

        var tableWidth = 0.0;
        for (var w : colWidths) tableWidth += w;
        var tableHeight = 0.0;
        for (var h : rowHeights) tableHeight += h;

        return new SheetLayout(fontScale, boxed(colWidths), boxed(rowHeights), tableWidth, tableHeight);
    }

    private static List<Double> boxed(double[] values) {
        var list = new ArrayList<Double>(values.length);
        for (var v : values) list.add(v);
        return list;
    }

    private static double naturalColWidth(List<RowSnapshot> rows, int col, double fontScale, MergeIndex mergeIndex) {
        var maxNeeded = LayoutConstants.MIN_COL_WIDTH_PT;
        for (int r = 0; r < rows.size(); r++) {
            var pos = new CellPos(r, col);
            if (mergeIndex.isCovered(pos) || mergeIndex.isWideAnchor(pos)) continue;
            var cell = rows.get(r).cellAt(col);
            if (cell.isEmpty()) continue;
            var baseFont = cell.fontSize() > 0 ? cell.fontSize() : LayoutConstants.DEFAULT_FONT_PT;
            var charFactor = LayoutConstants.CHAR_WIDTH_FACTOR * (cell.bold() ? 1.1 : 1.0);
            var width = cell.text().length() * (baseFont * fontScale) * charFactor + LayoutConstants.CELL_PADDING_PT;
            maxNeeded = Math.max(maxNeeded, width);
        }
        return maxNeeded;
    }

    private static double naturalRowHeight(RowSnapshot row, int rowIdx, double fontScale, MergeIndex mergeIndex) {
        var maxNeeded = LayoutConstants.MIN_ROW_HEIGHT_PT;
        for (int c = 0; c < row.cells().size(); c++) {
            var pos = new CellPos(rowIdx, c);
            if (mergeIndex.isCovered(pos) || mergeIndex.isTallAnchor(pos)) continue;
            var cell = row.cellAt(c);
            if (cell.isEmpty()) continue;
            var baseFont = cell.fontSize() > 0 ? cell.fontSize() : LayoutConstants.DEFAULT_FONT_PT;
            var needed = (baseFont * fontScale) * LayoutConstants.LINE_HEIGHT_FACTOR + LayoutConstants.ROW_PADDING_PT;
            maxNeeded = Math.max(maxNeeded, needed);
        }
        return maxNeeded;
    }
}
