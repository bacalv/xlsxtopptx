package xlsxtopptx;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

/** Reads one POI {@link Sheet} into a POI-independent {@link SheetSnapshot}. */
@Slf4j
public final class SheetReader {
    private SheetReader() {}

    public static SheetSnapshot read(Sheet sheet, int sheetFirstRow) {
        var lastRow = sheet.getLastRowNum();
        var lastCol = usedLastColumn(sheet, sheetFirstRow, lastRow);
        return read(sheet, sheetFirstRow, lastRow, 0, lastCol);
    }

    /** The sheet's full used area: every row from its first to its last, and every column up to
     *  the rightmost used one. */
    public static CellRangeAddress usedBounds(Sheet sheet) {
        var firstRow = Math.max(0, sheet.getFirstRowNum());
        var lastRow = sheet.getLastRowNum();
        var lastCol = usedLastColumn(sheet, firstRow, lastRow);
        return new CellRangeAddress(firstRow, lastRow, 0, lastCol);
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
        var themes = sheet.getWorkbook() instanceof XSSFWorkbook xwb ? xwb.getTheme() : null;
        var cfEvaluator = new ConditionalFormattingEvaluator(sheet.getWorkbook(), (WorkbookEvaluatorProvider) evaluator);

        var result = new ArrayList<RowSnapshot>();
        for (int r = firstRow; r <= lastRow; r++) {
            var row = sheet.getRow(r);
            var cells = new ArrayList<CellSnapshot>();
            for (int c = firstCol; c <= lastCol; c++) {
                var cell = row != null ? row.getCell(c) : null;
                cells.add(readCell(cell, evaluator, formatter, themes, cfEvaluator));
            }
            result.add(new RowSnapshot(cells));
        }
        return result;
    }

    private static CellSnapshot readCell(
        Cell cell, FormulaEvaluator evaluator, DataFormatter formatter, ThemesTable themes,
        ConditionalFormattingEvaluator cfEvaluator
    ) {
        if (cell == null) return CellSnapshot.empty();

        var value = readCellValue(cell, evaluator, formatter);
        var text = value.text();

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
            ? toAwtColor(xStyle.getFillForegroundColorColor(), themes)
            : null;
        var cfFill = conditionalFill(cell, cfEvaluator, themes);
        if (cfFill != null) fill = cfFill;

        var borders = new CellBorders(
            new BorderEdgeStyle(xStyle.getBorderTop(), toAwtColor(xStyle.getTopBorderXSSFColor(), themes)),
            new BorderEdgeStyle(xStyle.getBorderBottom(), toAwtColor(xStyle.getBottomBorderXSSFColor(), themes)),
            new BorderEdgeStyle(xStyle.getBorderLeft(), toAwtColor(xStyle.getLeftBorderXSSFColor(), themes)),
            new BorderEdgeStyle(xStyle.getBorderRight(), toAwtColor(xStyle.getRightBorderXSSFColor(), themes))
        );

        var font = xStyle.getFont();
        var fontColor = toAwtColor(font.getXSSFColor(), themes);

        return new CellSnapshot(
            text,
            fill,
            fontColor != null ? fontColor : LayoutConstants.DEFAULT_FONT_COLOR,
            font.getBold(),
            font.getItalic(),
            font.getFontHeightInPoints(),
            font.getFontName(),
            mapAlign(xStyle.getAlignment(), value.numericGeneral()),
            borders
        );
    }

    private record CellValue(String text, boolean numericGeneral) {}

    /** A cell's displayed text, and whether it should right-align as a number under Excel's
     *  default "General" alignment. Falls back to the formula's last-calculated (cached) result --
     *  skipping formula parsing/evaluation entirely -- when POI can't resolve a shared formula's
     *  master cell: a known POI limitation seen on real files after Excel's "Break Links" strips
     *  the formula text from a shared-formula group's master cell but leaves the other cells in
     *  that group still pointing at it (POI throws "Master cell of a shared formula with sid=... was
     *  not found"). Falls back to blank if even the cached result can't be read. */
    private static CellValue readCellValue(Cell cell, FormulaEvaluator evaluator, DataFormatter formatter) {
        try {
            return new CellValue(formatter.formatCellValue(cell, evaluator), isNumericGeneral(cell, evaluator));
        } catch (RuntimeException e) {
            var cause = sharedFormulaMasterMissingCause(e);
            if (cause == null) throw e;

            var fallback = cachedFormulaResult(cell, formatter);
            log.warn("{}: formula evaluation failed ({}) -- likely a shared formula elsewhere in the workbook "
                    + "broken by breaking external links in Excel; falling back to {}",
                cellReference(cell), cause.getMessage(), fallback.text().isEmpty() ? "blank" : "its last cached value");
            return fallback;
        }
    }

    /** Walks {@code e}'s cause chain for POI's "shared formula master not found" failure and
     *  returns the throwable that actually reports it, or {@code null} if it's not in there
     *  anywhere. This surfaces wrapped in an outer "Failed to evaluate cell: ..." exception
     *  whenever the cell being read merely *references* some other cell -- on any sheet -- whose
     *  own shared formula group is the one that's actually broken, so checking only {@code
     *  e.getMessage()} misses it. */
    private static Throwable sharedFormulaMasterMissingCause(Throwable e) {
        for (var t = e; t != null; t = t.getCause()) {
            if (t.getMessage() != null && t.getMessage().startsWith("Master cell of a shared formula")) {
                return t;
            }
        }
        return null;
    }

    private static String cellReference(Cell cell) {
        return cell.getSheet().getSheetName() + "!"
            + new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString();
    }

    /** Reads a formula cell's last-cached result directly off its cached result type/value,
     *  bypassing formula parsing entirely -- unlike {@link #isNumericGeneral}, which needs the
     *  formula evaluator (and so would hit the same missing-master failure this exists to work
     *  around). */
    private static CellValue cachedFormulaResult(Cell cell, DataFormatter formatter) {
        try {
            // A formula cell with no cached <v> at all reports its numeric value as 0.0 rather
            // than failing -- indistinguishable from a genuine cached zero unless checked directly.
            if (cell instanceof XSSFCell xssfCell && !xssfCell.getCTCell().isSetV()) {
                return new CellValue("", false);
            }

            var style = cell.getCellStyle();
            return switch (cell.getCachedFormulaResultType()) {
                case NUMERIC -> new CellValue(
                    formatter.formatRawCellContents(cell.getNumericCellValue(), style.getDataFormat(), style.getDataFormatString()),
                    true);
                case STRING -> new CellValue(cell.getStringCellValue(), false);
                case BOOLEAN -> new CellValue(cell.getBooleanCellValue() ? "TRUE" : "FALSE", false);
                default -> new CellValue("", false);
            };
        } catch (RuntimeException e) {
            return new CellValue("", false);
        }
    }

    /** Excel's default "General" horizontal alignment right-aligns numbers/dates and left-aligns
     *  text -- it isn't simply "left". A cell with no explicit alignment override reports
     *  HorizontalAlignment.GENERAL, so this resolves what that actually renders as by checking
     *  the cell's (possibly formula-evaluated) value type. */
    private static boolean isNumericGeneral(Cell cell, FormulaEvaluator evaluator) {
        var type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = evaluator.evaluateFormulaCell(cell);
        }
        return type == CellType.NUMERIC;
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
            if (row != null) {
                result.addAll(centerAcrossRunsInRow(row, r - firstRow, firstCol, lastCol));
            }
        }
        return result;
    }

    /** Finds runs of 2+ adjacent center-across-selection cells within a single row, as
     *  {@link MergeRegion}s relative to {@code firstCol} on {@code localRow}. */
    private static List<MergeRegion> centerAcrossRunsInRow(Row row, int localRow, int firstCol, int lastCol) {
        var runs = new ArrayList<MergeRegion>();
        var runStart = -1;
        // Scans one column past lastCol as a sentinel, so a run ending exactly at lastCol still
        // gets flushed into a MergeRegion instead of being silently dropped.
        for (int c = firstCol; c <= lastCol + 1; c++) {
            if (isCenterAcrossSelection(row, c, lastCol)) {
                runStart = runStart == -1 ? c : runStart;
            } else {
                runStart = flushRun(runs, runStart, localRow, firstCol, c);
            }
        }
        return runs;
    }

    private static boolean isCenterAcrossSelection(Row row, int col, int lastCol) {
        if (col > lastCol) return false;
        var cell = row.getCell(col);
        return cell != null && cell.getCellStyle().getAlignment() == HorizontalAlignment.CENTER_SELECTION;
    }

    /** Records the run starting at {@code runStart} (if any) as a {@link MergeRegion} ending at
     *  column {@code c - 1}, but only when it spans 2+ cells. Always returns -1, resetting the run
     *  for the next column. */
    private static int flushRun(List<MergeRegion> runs, int runStart, int localRow, int firstCol, int c) {
        if (runStart != -1 && c - 1 > runStart) {
            runs.add(new MergeRegion(localRow, localRow, runStart - firstCol, c - 1 - firstCol));
        }
        return -1;
    }

    // ---------- conditional formatting ----------

    /** The fill color of the highest-priority conditional-formatting rule whose condition is
     *  currently true for {@code cell} (Excel evaluates rules in priority order and the
     *  highest-priority one with a fill wins), or {@code null} if no applicable rule defines one. */
    private static Color conditionalFill(Cell cell, ConditionalFormattingEvaluator cfEvaluator, ThemesTable themes) {
        for (var match : cfEvaluator.getConditionalFormattingForCell(cell)) {
            var color = cfFillColor(match.getRule().getPatternFormatting(), themes);
            if (color != null) return color;
        }
        return null;
    }

    /** Conditional-formatting fills are stored as differential ("dxf") formatting records, where --
     *  unlike a normal cell style's {@code <fill>} -- a solid fill's color is written to {@code
     *  bgColor} rather than {@code fgColor}. This is a documented OOXML quirk specific to dxf
     *  records, and matches how POI's own API writes conditional-formatting fills
     *  ({@code PatternFormatting#setFillBackgroundColor}). */
    private static Color cfFillColor(PatternFormatting pattern, ThemesTable themes) {
        if (pattern == null || pattern.getFillPattern() == PatternFormatting.NO_FILL) return null;
        var color = pattern.getFillBackgroundColorColor();
        if (color == null) color = pattern.getFillForegroundColorColor();
        return color instanceof XSSFColor xColor ? toAwtColor(xColor, themes) : null;
    }

    // ---------- small conversions ----------

    /** Colors picked from Excel's "Theme Colors" palette (as opposed to "Standard Colors" or a
     *  custom RGB) are stored as a theme index plus an optional tint -- e.g. the light banded-row
     *  blue in a table is typically "Blue, Accent 1, Lighter 60%", not an RGB value at all.
     *  {@link XSSFColor#getRGB()} only ever returns an explicitly-stored RGB, so a themed color
     *  resolves to null (dropping the fill/border/font color entirely) unless it's first resolved
     *  against the workbook's theme, and the tint reapplied on top of that base color. */
    private static Color toAwtColor(XSSFColor color, ThemesTable themes) {
        if (color == null) return null;
        if (themes != null) themes.inheritFromThemeAsRequired(color);
        var rgb = color.getRGBWithTint();
        if (rgb == null) rgb = color.getRGB();
        if (rgb == null) return null;
        return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    private static TextAlign mapAlign(HorizontalAlignment alignment, boolean numericGeneral) {
        return switch (alignment) {
            case null -> numericGeneral ? TextAlign.RIGHT : TextAlign.LEFT;
            case GENERAL -> numericGeneral ? TextAlign.RIGHT : TextAlign.LEFT;
            case CENTER, CENTER_SELECTION -> TextAlign.CENTER;
            case RIGHT -> TextAlign.RIGHT;
            case JUSTIFY -> TextAlign.JUSTIFY;
            default -> TextAlign.LEFT;
        };
    }
}
