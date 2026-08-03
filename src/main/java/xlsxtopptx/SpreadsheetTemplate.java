package xlsxtopptx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
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
 *
 * <p>Separately, the sheet may contain sheet-wide {@code $NAME} tokens anywhere in a string cell
 * (e.g. a title cell reading {@code "As of $REPORT_DATE"}), each resolved once via {@link
 * #placeholder} rather than once per output row. Every {@code $NAME} token found in the sheet
 * must have a matching {@link #placeholder} call and vice versa -- an unknown token is almost
 * always a typo in the template, and an unused one is almost always a stale or misspelled call
 * left behind after the template changed.
 *
 * <p>{@link #replaceRange} covers a third, unrelated need: overwriting a fixed-size block of
 * cells (on any sheet, not just the marker sheet) with literal values, in place, with no row
 * growth or shrinkage -- e.g. a small KPI grid that's always exactly 3x2. Existing cell styles
 * are kept; only values change, and a formula cell is overwritten with a literal the same way
 * typing over it in Excel would. {@link #render} applies every registered {@code replaceRange}
 * before the row-expansion described above runs, so a range sitting below the marker block (e.g.
 * a footer table after a grand-total row) is carried along by that same row-expansion's {@code
 * shiftRows} call like any other trailing content, rather than needing its address computed in
 * post-shift coordinates. A range overlapping the marker block itself is rejected, since that
 * region is exclusively owned by row-expansion. Calling {@link #render} with no rows is allowed
 * as long as at least one {@code replaceRange} was registered -- row-expansion (and its
 * requirement that the sheet contain {@code %_<type>} markers) simply doesn't run.
 */
public final class SpreadsheetTemplate {
    private static final Pattern TYPE_MARKER = Pattern.compile("%_(?!\\d+$)(.+)");
    private static final Pattern PLACEHOLDER_MARKER = Pattern.compile("%_(\\d+)");
    private static final Pattern PLACEHOLDER_TOKEN = Pattern.compile("\\$([A-Za-z_][A-Za-z0-9_]*)");

    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;
    private final Map<String, String> placeholders = new HashMap<>();
    private final List<RangeReplacement> rangeReplacements = new ArrayList<>();

    private SpreadsheetTemplate(XSSFWorkbook workbook) {
        this.workbook = workbook;
        this.sheet = workbook.getSheetAt(0);
    }

    public static SpreadsheetTemplate of(InputStream templateInput) throws IOException {
        return new SpreadsheetTemplate(new XSSFWorkbook(templateInput));
    }

    /** Registers a sheet-wide {@code $name} token (case as given, e.g. {@code "REPORT_DATE"} for
     *  {@code $REPORT_DATE}) to be replaced by {@code value} everywhere it appears in the
     *  template's string cells. Must be called before {@link #render}. */
    public SpreadsheetTemplate placeholder(String name, Object value) {
        placeholders.put(name, String.valueOf(value));
        return this;
    }

    /** Registers a fixed-size block of {@code values} (row-major, one inner list per row) to
     *  overwrite in place at {@code a1Range} (e.g. {@code "A1:C5"}) on {@code sheetName} --
     *  {@code values} must match {@code a1Range}'s row/column count exactly. See the class
     *  Javadoc for how this interacts with row-expansion. Must be called before {@link #render}. */
    public SpreadsheetTemplate replaceRange(String sheetName, String a1Range, List<List<Object>> values) {
        rangeReplacements.add(new RangeReplacement(sheetName, CellRangeAddress.valueOf(a1Range), values));
        return this;
    }

    public byte[] render(List<SpreadsheetRow> rows) throws IOException {
        if (rows.isEmpty() && rangeReplacements.isEmpty()) {
            throw new IllegalArgumentException("rows must not be empty unless replaceRange(...) has been called");
        }

        Map<String, RowPrototype> prototypes = Map.of();
        var originalFirst = -1;
        var originalLast = -1;
        if (!rows.isEmpty()) {
            var found = findRowPrototypes();
            prototypes = found;
            var unknownTypes = rows.stream().map(SpreadsheetRow::type).filter(t -> !found.containsKey(t)).toList();
            if (!unknownTypes.isEmpty()) {
                throw new IllegalArgumentException(
                    "Unknown row type(s): " + unknownTypes + " -- known types: " + prototypes.keySet());
            }
            originalFirst = prototypes.values().stream().mapToInt(p -> p.rowIndex).min().orElseThrow();
            originalLast = prototypes.values().stream().mapToInt(p -> p.rowIndex).max().orElseThrow();
        }

        // Applied before row-expansion below so that a registered range sitting below the marker
        // block is written at its original template address, then carried along -- like any other
        // trailing content -- when shiftRows moves that address to make room for the actual output
        // row count.
        applyRangeReplacements(originalFirst, originalLast);

        if (!rows.isEmpty()) {
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
        }

        substitutePlaceholders();

        workbook.setForceFormulaRecalculation(true);
        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        var out = new ByteArrayOutputStream();
        workbook.write(out);
        return out.toByteArray();
    }

    /** Writes every registered {@link #replaceRange} block. {@code markerBlockFirstRow}/{@code
     *  markerBlockLastRow} are the row-expansion marker block's bounds on {@link #sheet} (0-based,
     *  inclusive), or {@code -1}/{@code -1} if {@link #render} was called with no rows -- in which
     *  case no marker block exists on this sheet to guard against. */
    private void applyRangeReplacements(int markerBlockFirstRow, int markerBlockLastRow) {
        for (var replacement : rangeReplacements) {
            applyRangeReplacement(replacement, markerBlockFirstRow, markerBlockLastRow);
        }
    }

    private void applyRangeReplacement(RangeReplacement replacement, int markerBlockFirstRow, int markerBlockLastRow) {
        var targetSheet = workbook.getSheet(replacement.sheetName());
        if (targetSheet == null) {
            throw new IllegalArgumentException("replaceRange: no such sheet \"" + replacement.sheetName() + "\"");
        }

        var range = replacement.range();
        if (targetSheet == sheet && range.getFirstRow() <= markerBlockLastRow && range.getLastRow() >= markerBlockFirstRow) {
            throw new IllegalArgumentException("replaceRange(\"" + replacement.sheetName() + "\", \"" + range.formatAsString()
                + "\") overlaps the row-expansion marker block (rows " + (markerBlockFirstRow + 1) + "-" + (markerBlockLastRow + 1)
                + ") -- that region is rewritten by row-expansion regardless of this call");
        }

        var expectedRows = range.getLastRow() - range.getFirstRow() + 1;
        var expectedCols = range.getLastColumn() - range.getFirstColumn() + 1;
        var values = replacement.values();
        if (values.size() != expectedRows || values.stream().anyMatch(row -> row.size() != expectedCols)) {
            throw new IllegalArgumentException("replaceRange(\"" + replacement.sheetName() + "\", \"" + range.formatAsString()
                + "\") needs a " + expectedRows + "x" + expectedCols + " grid of values");
        }

        for (int r = 0; r < expectedRows; r++) {
            var row = targetSheet.getRow(range.getFirstRow() + r);
            if (row == null) {
                row = targetSheet.createRow(range.getFirstRow() + r);
            }
            for (int c = 0; c < expectedCols; c++) {
                setValue(freshValueCell(row, range.getFirstColumn() + c), values.get(r).get(c));
            }
        }
    }

    /** The cell at {@code col} in {@code row}, ready to take a literal value -- creating it if
     *  absent. {@code Cell.setCellValue()} on an existing {@code FORMULA} cell only updates its
     *  cached result and deliberately leaves the cell type as {@code FORMULA} (per its own
     *  Javadoc), so such a cell is recreated from scratch first, the same way {@link
     *  #deshareFormulasInRow} resets a cell's type -- keeping its style. */
    private static Cell freshValueCell(Row row, int col) {
        var existing = row.getCell(col);
        if (existing == null) {
            return row.createCell(col);
        }
        if (existing.getCellType() != CellType.FORMULA) {
            return existing;
        }

        var style = existing.getCellStyle();
        row.removeCell(existing);
        var cell = row.createCell(col);
        cell.setCellStyle(style);
        return cell;
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

    /** Replaces every {@code $NAME} token in the sheet's string cells with its registered {@link
     *  #placeholder} value, after checking that the tokens actually present in the sheet exactly
     *  match the placeholders that were registered. Runs after rows are written so it also covers
     *  any token that happens to land in freshly written row data, though it's meant for static
     *  cells (titles, headers) outside the marker block. */
    private void substitutePlaceholders() {
        var foundTokens = new HashSet<String>();
        for (var row : sheet) {
            for (var cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    collectTokens(cell.getStringCellValue(), foundTokens);
                }
            }
        }

        var unknownTokens = foundTokens.stream().filter(t -> !placeholders.containsKey(t)).toList();
        if (!unknownTokens.isEmpty()) {
            throw new IllegalArgumentException(
                "Unknown placeholder(s) in template: " + unknownTokens + " -- registered placeholder(s): "
                    + placeholders.keySet());
        }
        var unusedPlaceholders = placeholders.keySet().stream().filter(p -> !foundTokens.contains(p)).toList();
        if (!unusedPlaceholders.isEmpty()) {
            throw new IllegalArgumentException("Placeholder(s) registered but not found in template: "
                + unusedPlaceholders);
        }

        for (var row : sheet) {
            for (var cell : row) {
                if (cell.getCellType() == CellType.STRING && PLACEHOLDER_TOKEN.matcher(cell.getStringCellValue()).find()) {
                    cell.setCellValue(replaceTokens(cell.getStringCellValue()));
                }
            }
        }
    }

    private static void collectTokens(String text, Set<String> into) {
        var matcher = PLACEHOLDER_TOKEN.matcher(text);
        while (matcher.find()) {
            into.add(matcher.group(1));
        }
    }

    private String replaceTokens(String text) {
        var matcher = PLACEHOLDER_TOKEN.matcher(text);
        var result = new StringBuilder();
        while (matcher.find()) {
            matcher.appendReplacement(result, Matcher.quoteReplacement(placeholders.get(matcher.group(1))));
        }
        matcher.appendTail(result);
        return result.toString();
    }

    private boolean referencesThisSheet(String reference) {
        var bangIndex = reference.indexOf('!');
        if (bangIndex < 0) return false;
        var sheetPart = reference.substring(0, bangIndex).replace("'", "");
        return sheetPart.equals(sheet.getSheetName());
    }

    /** One pending {@link #replaceRange} call, held until {@link #render} applies it. */
    private record RangeReplacement(String sheetName, CellRangeAddress range, List<List<Object>> values) {
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
