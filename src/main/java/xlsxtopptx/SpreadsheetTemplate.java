package xlsxtopptx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * Renders {@link SpreadsheetRow}s onto a "row template" .xlsx: a sheet whose first sheet carries
 * a handful of rows marked in their label column with {@code %_<type>} (e.g. {@code %_TEMPLATE
 * ROW BOLD}), each with {@code %_0}, {@code %_1}, ... placeholder cells and optionally its own
 * per-row formulas (e.g. {@code SUM(B4:F4)}).
 *
 * <p>{@link #render} replaces that marker block with one output row per {@code SpreadsheetRow},
 * cloning the matching marker row's cell styles and formulas -- formulas have their row
 * references shifted to the output row, the same way a spreadsheet application would when you
 * drag-fill a formula down. Rows below the marker block (e.g. a grand-total row) are moved to sit
 * directly after the output rows; for a formula there to keep working across a varying number of
 * output rows, it must reference the marker block via a named range (e.g. {@code
 * SUM(COL_B_VALUES)}) rather than a hardcoded cell range, since only such named ranges are grown
 * or shrunk to cover the actual output rows. Marker rows must be contiguous in the sheet.
 */
public final class SpreadsheetTemplate {
    private static final Pattern TYPE_MARKER = Pattern.compile("%_(?!\\d+$)(.+)");
    private static final Pattern PLACEHOLDER_MARKER = Pattern.compile("%_(\\d+)");

    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;

    private SpreadsheetTemplate(XSSFWorkbook workbook) {
        this.workbook = workbook;
        this.sheet = workbook.getSheetAt(0);
    }

    public static SpreadsheetTemplate of(InputStream templateInput) throws IOException {
        return new SpreadsheetTemplate(new XSSFWorkbook(templateInput));
    }

    public byte[] render(List<SpreadsheetRow> rows) throws IOException {
        if (rows.isEmpty()) {
            throw new IllegalArgumentException("rows must not be empty");
        }

        var prototypes = findRowPrototypes();
        var unknownTypes = rows.stream().map(SpreadsheetRow::type).filter(t -> !prototypes.containsKey(t)).toList();
        if (!unknownTypes.isEmpty()) {
            throw new IllegalArgumentException(
                "Unknown row type(s): " + unknownTypes + " -- known types: " + prototypes.keySet());
        }

        var originalFirst = prototypes.values().stream().mapToInt(p -> p.rowIndex).min().orElseThrow();
        var originalLast = prototypes.values().stream().mapToInt(p -> p.rowIndex).max().orElseThrow();
        var newLast = originalFirst + rows.size() - 1;
        var delta = newLast - originalLast;

        // De-shares formulas below the marker block regardless of delta: a human-authored
        // trailing row (e.g. a grand total someone dragged a formula across in Excel/Sheets) can
        // carry stale "shared formula" grouping even when row count -- and so delta -- ends up
        // zero, which is enough on its own to make POI misread it later (see deshareFormulas).
        if (originalLast < sheet.getLastRowNum()) {
            deshareFormulas(originalLast + 1, sheet.getLastRowNum());
        }

        // Moves everything below the marker block (e.g. a grand-total row) so it sits directly
        // after the output rows. Rows within the marker block itself are untouched by this --
        // they're overwritten below regardless of delta's sign, since every prototype needed was
        // already captured above.
        if (delta != 0 && originalLast < sheet.getLastRowNum()) {
            sheet.shiftRows(originalLast + 1, sheet.getLastRowNum(), delta);
        }

        for (int i = 0; i < rows.size(); i++) {
            writeRow(originalFirst + i, prototypes.get(rows.get(i).type()), rows.get(i));
        }

        growNamedRanges(originalLast, newLast);
        trimTrailingEmptyRows(newLast);

        workbook.setForceFormulaRecalculation(true);
        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        var out = new ByteArrayOutputStream();
        workbook.write(out);
        return out.toByteArray();
    }

    /** Scans the sheet for {@code %_<type>} marker cells and captures each marked row's cell
     *  styles, placeholder positions, and formulas before anything is modified. */
    private Map<String, RowPrototype> findRowPrototypes() {
        var prototypes = new HashMap<String, RowPrototype>();
        for (var row : sheet) {
            for (var cell : row) {
                var type = typeMarkerOf(cell);
                if (type != null) {
                    captureRowPrototype(prototypes, type, row, cell);
                }
            }
        }
        if (prototypes.isEmpty()) {
            throw new IllegalStateException("No \"%_<type>\" row markers found in " + sheet.getSheetName());
        }
        return prototypes;
    }

    private static void captureRowPrototype(Map<String, RowPrototype> prototypes, String type, Row row, Cell cell) {
        // computeIfAbsent's mapping function only runs the first time "type" is seen -- wasNew
        // tracks whether that happened just now, so a later cell reusing the same type marker
        // (same row or a different one) is caught as a duplicate instead of silently ignored.
        var wasNew = new boolean[1];
        prototypes.computeIfAbsent(type, t -> {
            wasNew[0] = true;
            return RowPrototype.capture(row, cell.getColumnIndex());
        });
        if (!wasNew[0]) {
            throw new IllegalStateException("Duplicate row type marker: %_" + type);
        }
    }

    /** The {@code <type>} out of a {@code %_<type>} marker cell, or null if {@code cell} isn't
     *  one (not a string cell, or its text doesn't match the marker pattern). */
    private static String typeMarkerOf(Cell cell) {
        if (cell.getCellType() != CellType.STRING) return null;
        var matcher = TYPE_MARKER.matcher(cell.getStringCellValue());
        return matcher.matches() ? matcher.group(1) : null;
    }

    private void writeRow(int outputRowIndex, RowPrototype prototype, SpreadsheetRow data) {
        var expectedPlaceholders = prototype.placeholderColumns.size();
        if (data.values().size() != expectedPlaceholders) {
            throw new IllegalArgumentException("Row type \"" + data.type() + "\" needs " + expectedPlaceholders
                + " value(s) but got " + data.values().size());
        }

        var existing = sheet.getRow(outputRowIndex);
        if (existing != null) {
            sheet.removeRow(existing);
        }
        var row = sheet.createRow(outputRowIndex);
        row.setHeightInPoints(prototype.height);

        for (var entry : prototype.stylesByColumn.entrySet()) {
            var col = entry.getKey();
            var cell = row.createCell(col);
            cell.setCellStyle(entry.getValue());

            if (col == prototype.labelColumnIndex) {
                cell.setCellValue(data.label());
            } else if (prototype.placeholderColumns.containsKey(col)) {
                setValue(cell, data.values().get(prototype.placeholderColumns.get(col)));
            } else if (prototype.formulasByColumn.containsKey(col)) {
                cell.setCellFormula(
                    shiftFormula(prototype.formulasByColumn.get(col), prototype.rowNumber1Based, outputRowIndex + 1));
            }
        }
    }

    private static void setValue(Cell cell, Object value) {
        if (value instanceof Number number) {
            cell.setCellValue(number.doubleValue());
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }

    /** Rewrites cell references in {@code formula} from one row to another, the same way a
     *  spreadsheet application shifts a formula's row-relative references when you copy it to a
     *  different row -- e.g. {@code SUM(B4:F4)} copied from row 4 to row 6 becomes
     *  {@code SUM(B6:F6)}. */
    private static String shiftFormula(String formula, int fromRow1Based, int toRow1Based) {
        var pattern = Pattern.compile("([A-Z]+)" + fromRow1Based + "\\b");
        return pattern.matcher(formula).replaceAll("$1" + toRow1Based);
    }

    /** Grows or shrinks any workbook-level named range that ends exactly at the marker block's
     *  last row on this sheet, so formulas referencing it (e.g. {@code SUM(COL_B_VALUES)}) cover
     *  the actual output rows instead of the original template block. Named ranges on other
     *  sheets, or not ending at the marker block, are left untouched. */
    private void growNamedRanges(int originalLastRowIndex, int newLastRowIndex) {
        var oldRowNumber = originalLastRowIndex + 1;
        var newRowNumber = newLastRowIndex + 1;
        var rowEndPattern = Pattern.compile("(\\$[A-Z]+\\$)" + oldRowNumber + "$");

        for (var name : workbook.getAllNames()) {
            growNamedRange(name, rowEndPattern, newRowNumber);
        }
    }

    private void growNamedRange(XSSFName name, Pattern rowEndPattern, int newRowNumber) {
        var reference = name.getRefersToFormula();
        if (!referencesThisSheet(reference)) return;

        var matcher = rowEndPattern.matcher(reference);
        if (matcher.find()) {
            name.setRefersToFormula(matcher.replaceFirst("$1" + newRowNumber));
        }
    }

    /** Recreates every formula cell in {@code [firstRowIndex, lastRowIndex]}, keeping its
     *  resolved text and style but dropping any Excel "shared formula" grouping -- a real risk
     *  here since a human-authored template's trailing rows (e.g. a grand-total row someone built
     *  by dragging a formula across columns in Excel/Sheets) are likely to carry one. POI's
     *  {@code Cell.getCellFormula()} resolves a shared-formula sibling by shifting the group
     *  master's formula across their column/row offset, which for a formula with no
     *  column-relative token to shift (like a named-range reference) collapses every sibling onto
     *  the master's exact text -- even though, as here, the sibling's own stored text is both
     *  present and correct. This misreading happens on any load, independent of whether
     *  {@code Sheet.shiftRows} is subsequently used to move the row. */
    private void deshareFormulas(int firstRowIndex, int lastRowIndex) {
        for (int r = firstRowIndex; r <= lastRowIndex; r++) {
            deshareFormulasInRow(sheet.getRow(r));
        }
    }

    private static void deshareFormulasInRow(Row row) {
        if (row == null) return;

        var formulaCells = new ArrayList<Cell>();
        for (var cell : row) {
            if (cell.getCellType() == CellType.FORMULA) {
                formulaCells.add(cell);
            }
        }

        for (var cell : formulaCells) {
            var col = cell.getColumnIndex();
            var style = cell.getCellStyle();
            var formula = rawFormulaText(cell);

            row.removeCell(cell);
            var newCell = row.createCell(col);
            newCell.setCellStyle(style);
            newCell.setCellFormula(formula);
        }
    }

    /** The cell's own stored formula text, bypassing shared-formula master derivation.
     *  {@code Cell.getCellFormula()} resolves a shared-formula sibling by shifting the group
     *  master's formula across the column/row offset between them -- which silently discards
     *  whatever text the sibling cell itself actually stores whenever that offset has nothing to
     *  shift (as with a named-range reference), even when, as here, that stored text is both
     *  present and correct. Falls back to {@code getCellFormula()} for a sibling that (as real
     *  shared-formula compression does) stores no text of its own. */
    private static String rawFormulaText(Cell cell) {
        var raw = ((org.apache.poi.xssf.usermodel.XSSFCell) cell).getCTCell().getF().getStringValue();
        return raw == null || raw.isEmpty() ? cell.getCellFormula() : raw;
    }

    /** Drops genuinely empty rows (no cells at all) off the bottom of the sheet, stopping at the
     *  first non-empty one -- a template hand-authored in Excel/Sheets can carry a large tail of
     *  rows with nothing but row-height formatting (no cell data), which would otherwise inflate
     *  the sheet's apparent "used range" for any downstream consumer (e.g. this project's own
     *  SpreadsheetToPptx) far past the last row this method actually wrote. Never trims into or
     *  above the output rows themselves. */
    private void trimTrailingEmptyRows(int newLastRowIndex) {
        for (int r = sheet.getLastRowNum(); r > newLastRowIndex && isEmptyRow(sheet.getRow(r)); r--) {
            var row = sheet.getRow(r);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private static boolean isEmptyRow(Row row) {
        return row == null || row.getPhysicalNumberOfCells() == 0;
    }

    private boolean referencesThisSheet(String reference) {
        var bangIndex = reference.indexOf('!');
        if (bangIndex < 0) return false;
        var sheetPart = reference.substring(0, bangIndex).replace("'", "");
        return sheetPart.equals(sheet.getSheetName());
    }

    /** A captured, in-memory snapshot of one marker row, taken before any row in the sheet is
     *  modified -- so it stays valid even if the same type is used for several output rows, or
     *  the marker row's own position gets overwritten while writing output. */
    private static final class RowPrototype {
        final int rowIndex;
        final int rowNumber1Based;
        final float height;
        final int labelColumnIndex;
        final Map<Integer, org.apache.poi.ss.usermodel.CellStyle> stylesByColumn = new HashMap<>();
        final Map<Integer, Integer> placeholderColumns = new HashMap<>();
        final Map<Integer, String> formulasByColumn = new HashMap<>();

        private RowPrototype(int rowIndex, float height, int labelColumnIndex) {
            this.rowIndex = rowIndex;
            this.rowNumber1Based = rowIndex + 1;
            this.height = height;
            this.labelColumnIndex = labelColumnIndex;
        }

        static RowPrototype capture(Row row, int labelColumnIndex) {
            var prototype = new RowPrototype(row.getRowNum(), row.getHeightInPoints(), labelColumnIndex);
            for (var cell : row) {
                var col = cell.getColumnIndex();
                prototype.stylesByColumn.put(col, cell.getCellStyle());

                if (cell.getCellType() == CellType.FORMULA) {
                    prototype.formulasByColumn.put(col, cell.getCellFormula());
                } else if (col != labelColumnIndex && cell.getCellType() == CellType.STRING) {
                    var matcher = PLACEHOLDER_MARKER.matcher(cell.getStringCellValue());
                    if (matcher.matches()) {
                        prototype.placeholderColumns.put(col, Integer.parseInt(matcher.group(1)));
                    }
                }
            }
            return prototype;
        }
    }
}
