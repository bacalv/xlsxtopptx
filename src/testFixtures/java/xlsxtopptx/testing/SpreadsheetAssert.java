package xlsxtopptx.testing;

import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.UncheckedIOException;

/**
 * Asserts that two .xlsx workbooks are semantically equal: same sheets, cell text/values,
 * formulas, and styling (fonts, fills, borders, alignment, merges) -- while ignoring incidental
 * byte-level differences between two writer runs, such as timestamps, calc chain, or XML
 * attribute ordering. Regenerating "the same" spreadsheet from the same inputs at a different time
 * will produce different bytes but should still compare equal here.
 *
 * <p>Framework-agnostic: throws a plain {@link AssertionError} on failure, so it works the same
 * under JUnit, TestNG, or any other runner.
 */
public final class SpreadsheetAssert {
    private SpreadsheetAssert() {}

    /** @throws AssertionError if the two workbooks are not semantically equal. */
    public static void assertEquals(byte[] expectedXlsx, byte[] actualXlsx) {
        assertEquals(null, expectedXlsx, actualXlsx);
    }

    /** @throws AssertionError if the two workbooks are not semantically equal. {@code message} is
     *  prefixed to the failure, same as JUnit's message-first assertion overloads. */
    public static void assertEquals(String message, byte[] expectedXlsx, byte[] actualXlsx) {
        try (var expected = WorkbookFactory.create(new ByteArrayInputStream(expectedXlsx));
             var actual = WorkbookFactory.create(new ByteArrayInputStream(actualXlsx))) {
            var diffs = SpreadsheetComparator.compare(expected, actual);
            DiffReport.failIfAny(message, "spreadsheets are not semantically equal", diffs);
        } catch (IOException e) {
            throw new UncheckedIOException("failed to read workbook bytes for comparison", e);
        }
    }
}
