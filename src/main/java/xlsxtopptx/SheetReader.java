package xlsxtopptx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

/** Reads one POI {@link Sheet} into a POI-independent {@link SheetSnapshot}. */
public final class SheetReader {
    private SheetReader() {}

    public static SheetSnapshot read(Sheet sheet, int sheetFirstRow) {
        var lastRow = sheet.getLastRowNum();
        var lastCol = usedLastColumn(sheet, sheetFirstRow, lastRow);
        return read(sheet, sheetFirstRow, lastRow, 0, lastCol);
    }

    /** The rightmost used column (inclusive, zero-based) across {@code [firstRow, lastRow]},
     *  accounting for merged regions whose non-anchor cells often have no backing Cell object at
     *  all (so a plain cell scan alone can undercount). Returns -1 if nothing is used. */
    public static int usedLastColumn(Sheet sheet, int firstRow, int lastRow) {
        var lastCol = -1;
        for (int r = firstRow; r <= lastRow; r++) {
            var row = sheet.getRow(r);
            if (row != null) lastCol = Math.max(lastCol, row.getLastCellNum() - 1);
        }
        for (var ra : sheet.getMergedRegions()) {
            if (ra.getLastRow() >= firstRow && ra.getFirstRow() <= lastRow) {
                lastCol = Math.max(lastCol, ra.getLastColumn());
            }
        }
        return lastCol;
    }

    /** Reads only the given inclusive, zero-based row/column bounds of {@code sheet}. */
    public static SheetSnapshot read(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        var rows = readRows(sheet, firstRow, lastRow, firstCol, lastCol);
        var numCols = Math.max(1, lastCol - firstCol + 1);
        var merges = new ArrayList<>(readMerges(sheet, firstRow, lastRow, firstCol, lastCol));
        merges.addAll(detectCenterAcrossSelectionRuns(sheet, firstRow, lastRow, firstCol, lastCol));
        return new SheetSnapshot(rows, merges, numCols);
    }

    // ---------- rows and cells ----------

    private static List<RowSnapshot> readRows(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        var evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        var formatter = new DataFormatter();

        var result = new ArrayList<RowSnapshot>();
        for (int r = firstRow; r <= lastRow; r++) {
            var row = sheet.getRow(r);
            var cells = new ArrayList<CellSnapshot>();
            for (int c = firstCol; c <= lastCol; c++) {
                var cell = row != null ? row.getCell(c) : null;
                cells.add(readCell(cell, evaluator, formatter));
            }
            result.add(new RowSnapshot(cells));
        }
        return result;
    }

    private static CellSnapshot readCell(Cell cell, FormulaEvaluator evaluator, DataFormatter formatter) {
        if (cell == null) return CellSnapshot.empty();

        var text = formatter.formatCellValue(cell, evaluator);

        var style = cell.getCellStyle();
        if (!(style instanceof XSSFCellStyle xStyle)) {
            return new CellSnapshot(text, null, LayoutConstants.DEFAULT_FONT_COLOR, false, false,
                LayoutConstants.DEFAULT_FONT_PT, null, TextAlign.LEFT, CellBorders.NONE);
        }

        // XSSFCellStyle.getBorder*() return NONE whenever the <xf> is missing an explicit
        // applyBorder="1" attribute -- it checks getApplyBorder() with no fallback. Many real
        // XLSX writers (e.g. openpyxl) define a genuine border on the <xf> but never emit that
        // flag, since Excel itself doesn't require it. POI already special-cases exactly this for
        // fills (treats a missing applyFill as "apply"), just not for borders, so we force it here
        // to match how every real spreadsheet app actually renders these files.
        xStyle.getCoreXf().setApplyBorder(true);

        var fill = xStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND
            ? toAwtColor(xStyle.getFillForegroundColorColor())
            : null;

        var borders = new CellBorders(
            new BorderEdgeStyle(xStyle.getBorderTop(), toAwtColor(xStyle.getTopBorderXSSFColor())),
            new BorderEdgeStyle(xStyle.getBorderBottom(), toAwtColor(xStyle.getBottomBorderXSSFColor())),
            new BorderEdgeStyle(xStyle.getBorderLeft(), toAwtColor(xStyle.getLeftBorderXSSFColor())),
            new BorderEdgeStyle(xStyle.getBorderRight(), toAwtColor(xStyle.getRightBorderXSSFColor()))
        );

        var font = xStyle.getFont();
        var fontColor = toAwtColor(font.getXSSFColor());

        return new CellSnapshot(
            text,
            fill,
            fontColor != null ? fontColor : LayoutConstants.DEFAULT_FONT_COLOR,
            font.getBold(),
            font.getItalic(),
            font.getFontHeightInPoints(),
            font.getFontName(),
            mapAlign(xStyle.getAlignment()),
            borders
        );
    }

    // ---------- merges ----------

    private static List<MergeRegion> readMerges(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        var merges = new ArrayList<MergeRegion>();
        for (var ra : sheet.getMergedRegions()) {
            if (ra.getLastRow() < firstRow || ra.getFirstRow() > lastRow
                || ra.getLastColumn() < firstCol || ra.getFirstColumn() > lastCol) {
                continue;
            }
            merges.add(new MergeRegion(
                Math.max(ra.getFirstRow(), firstRow) - firstRow, Math.min(ra.getLastRow(), lastRow) - firstRow,
                Math.max(ra.getFirstColumn(), firstCol) - firstCol, Math.min(ra.getLastColumn(), lastCol) - firstCol
            ));
        }
        return merges;
    }

    /** Finds runs of 2+ adjacent cells in the same row sharing HorizontalAlignment.CENTER_SELECTION
     *  ("Center Across Selection" in Excel) -- a look-alike-merge that getMergedRegions() won't report. */
    private static List<MergeRegion> detectCenterAcrossSelectionRuns(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        var result = new ArrayList<MergeRegion>();
        for (int r = firstRow; r <= lastRow; r++) {
            var row = sheet.getRow(r);
            if (row == null) continue;
            var runStart = -1;
            for (int c = firstCol; c <= lastCol + 1; c++) {
                var isCenterAcross = false;
                if (c <= lastCol) {
                    var cell = row.getCell(c);
                    if (cell != null) {
                        isCenterAcross = cell.getCellStyle().getAlignment() == HorizontalAlignment.CENTER_SELECTION;
                    }
                }
                if (isCenterAcross) {
                    if (runStart == -1) runStart = c;
                } else {
                    if (runStart != -1 && c - 1 > runStart) {
                        result.add(new MergeRegion(r - firstRow, r - firstRow, runStart - firstCol, c - 1 - firstCol));
                    }
                    runStart = -1;
                }
            }
        }
        return result;
    }

    // ---------- small conversions ----------

    private static Color toAwtColor(XSSFColor color) {
        if (color == null) return null;
        var rgb = color.getRGB();
        if (rgb == null) return null;
        return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    private static TextAlign mapAlign(HorizontalAlignment alignment) {
        return switch (alignment) {
            case null -> TextAlign.LEFT;
            case CENTER, CENTER_SELECTION -> TextAlign.CENTER;
            case RIGHT -> TextAlign.RIGHT;
            case JUSTIFY -> TextAlign.JUSTIFY;
            default -> TextAlign.LEFT;
        };
    }
}
