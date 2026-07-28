package xlsxtopptx.testing;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import xlsxtopptx.CellSnapshot;
import xlsxtopptx.MergeRegion;
import xlsxtopptx.RowSnapshot;
import xlsxtopptx.SheetReader;
import xlsxtopptx.SheetSnapshot;

import java.awt.Color;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/**
 * Compares two workbooks cell-by-cell via the library's own {@link SheetReader}, which already
 * normalizes each cell down to what it looks like rendered (text, resolved fill/font colors,
 * bold/italic, borders) rather than the raw XML -- exactly the incidental byte-level noise (calc
 * chain, XML attribute ordering, POI's own formatting quirks) that makes two writer runs of the
 * "same" spreadsheet differ at the byte level. Formula text is compared separately, since
 * SheetReader only captures each formula's evaluated display value.
 */
final class SpreadsheetComparator {
    private static final RowSnapshot EMPTY_ROW = new RowSnapshot(List.of());

    private SpreadsheetComparator() {}

    static List<String> compare(Workbook expected, Workbook actual) {
        var diffs = new ArrayList<String>();

        var expectedCount = expected.getNumberOfSheets();
        var actualCount = actual.getNumberOfSheets();
        if (expectedCount != actualCount) {
            diffs.add("sheet count differs: expected=%d actual=%d".formatted(expectedCount, actualCount));
        }

        var n = Math.min(expectedCount, actualCount);
        for (int i = 0; i < n; i++) {
            compareSheet(i, expected.getSheetAt(i), actual.getSheetAt(i), diffs);
        }
        return diffs;
    }

    private static void compareSheet(int index, Sheet expected, Sheet actual, List<String> diffs) {
        var prefix = "sheet[%d]".formatted(index);
        if (!expected.getSheetName().equals(actual.getSheetName())) {
            diffs.add("%s: name differs: expected=\"%s\" actual=\"%s\"".formatted(
                prefix, expected.getSheetName(), actual.getSheetName()));
        }
        prefix = prefix + " \"" + expected.getSheetName() + "\"";

        var expectedSnapshot = SheetReader.read(expected, 0);
        var actualSnapshot = SheetReader.read(actual, 0);

        compareMerges(prefix, expectedSnapshot.merges(), actualSnapshot.merges(), diffs);
        compareCells(prefix, expectedSnapshot, actualSnapshot, diffs);
        compareFormulas(prefix, expected, actual, diffs);
    }

    private static void compareMerges(String prefix, List<MergeRegion> expected, List<MergeRegion> actual, List<String> diffs) {
        var expectedSet = new LinkedHashSet<>(expected);
        var actualSet = new LinkedHashSet<>(actual);
        if (!expectedSet.equals(actualSet)) {
            diffs.add("%s: merged regions differ: expected=%s actual=%s".formatted(prefix, expectedSet, actualSet));
        }
    }

    private static void compareCells(String prefix, SheetSnapshot expected, SheetSnapshot actual, List<String> diffs) {
        var rowCount = Math.max(expected.rowCount(), actual.rowCount());
        var colCount = Math.max(expected.numCols(), actual.numCols());

        for (int r = 0; r < rowCount; r++) {
            var expectedRow = r < expected.rows().size() ? expected.rows().get(r) : EMPTY_ROW;
            var actualRow = r < actual.rows().size() ? actual.rows().get(r) : EMPTY_ROW;

            for (int c = 0; c < colCount; c++) {
                var expectedCell = expectedRow.cellAt(c);
                var actualCell = actualRow.cellAt(c);
                if (!expectedCell.equals(actualCell)) {
                    diffs.add(describeCellDiff(prefix, r, c, expectedCell, actualCell));
                }
            }
        }
    }

    private static String describeCellDiff(String prefix, int row, int col, CellSnapshot expected, CellSnapshot actual) {
        var fieldDiffs = new ArrayList<String>();
        if (!Objects.equals(expected.text(), actual.text())) {
            fieldDiffs.add("text: expected=\"%s\" actual=\"%s\"".formatted(expected.text(), actual.text()));
        }
        if (!Objects.equals(expected.fill(), actual.fill())) {
            fieldDiffs.add("fill: expected=%s actual=%s".formatted(colorStr(expected.fill()), colorStr(actual.fill())));
        }
        if (!Objects.equals(expected.fontColor(), actual.fontColor())) {
            fieldDiffs.add("fontColor: expected=%s actual=%s".formatted(colorStr(expected.fontColor()), colorStr(actual.fontColor())));
        }
        if (expected.bold() != actual.bold()) {
            fieldDiffs.add("bold: expected=%s actual=%s".formatted(expected.bold(), actual.bold()));
        }
        if (expected.italic() != actual.italic()) {
            fieldDiffs.add("italic: expected=%s actual=%s".formatted(expected.italic(), actual.italic()));
        }
        if (Double.compare(expected.fontSize(), actual.fontSize()) != 0) {
            fieldDiffs.add("fontSize: expected=%s actual=%s".formatted(expected.fontSize(), actual.fontSize()));
        }
        if (!Objects.equals(expected.fontName(), actual.fontName())) {
            fieldDiffs.add("fontName: expected=%s actual=%s".formatted(expected.fontName(), actual.fontName()));
        }
        if (expected.align() != actual.align()) {
            fieldDiffs.add("align: expected=%s actual=%s".formatted(expected.align(), actual.align()));
        }
        if (!Objects.equals(expected.borders(), actual.borders())) {
            fieldDiffs.add("borders: expected=%s actual=%s".formatted(expected.borders(), actual.borders()));
        }
        return "%s: cell[row=%d, col=%d]: %s".formatted(prefix, row, col, String.join(", ", fieldDiffs));
    }

    // SheetReader deliberately captures a cell's evaluated display text, not its formula text, so
    // two formulas that currently evaluate the same (e.g. "=1+2" vs "=3") wouldn't otherwise be
    // told apart -- this pass compares the raw formula strings on top of that.
    private static void compareFormulas(String prefix, Sheet expected, Sheet actual, List<String> diffs) {
        var lastRow = Math.max(expected.getLastRowNum(), actual.getLastRowNum());
        for (int r = 0; r <= lastRow; r++) {
            var expectedRow = expected.getRow(r);
            var actualRow = actual.getRow(r);
            var lastCol = Math.max(lastCellNum(expectedRow), lastCellNum(actualRow));

            for (int c = 0; c < lastCol; c++) {
                var expectedFormula = formulaOf(expectedRow, c);
                var actualFormula = formulaOf(actualRow, c);
                if (!Objects.equals(expectedFormula, actualFormula)) {
                    diffs.add("%s: cell[row=%d, col=%d]: formula: expected=%s actual=%s".formatted(
                        prefix, r, c, quote(expectedFormula), quote(actualFormula)));
                }
            }
        }
    }

    private static int lastCellNum(Row row) {
        return row == null ? 0 : Math.max(0, (int) row.getLastCellNum());
    }

    private static String formulaOf(Row row, int col) {
        if (row == null) return null;
        var cell = row.getCell(col);
        if (cell == null || cell.getCellType() != CellType.FORMULA) return null;
        return cell.getCellFormula();
    }

    private static String quote(String s) {
        return s == null ? "null" : "\"" + s + "\"";
    }

    private static String colorStr(Color c) {
        return c == null ? "null" : "#%02X%02X%02X".formatted(c.getRed(), c.getGreen(), c.getBlue());
    }
}
