package xlsxtopptx;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.time.Instant;
import java.util.Arrays;
import java.util.Date;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static xlsxtopptx.testing.SpreadsheetAssert.assertEquals;

/**
 * Exercises {@link xlsxtopptx.testing.SpreadsheetAssert} (in the testFixtures source set) the same
 * way a downstream service's end-to-end test would: regenerate the fixture workbook a second time
 * and confirm the assertion treats it as equal despite the raw bytes differing, then confirm a
 * genuine content difference is still caught.
 */
class SpreadsheetAssertTest {

    @Test
    void regeneratingTheSameWorkbookProducesDifferentBytesButComparesEqual() throws Exception {
        // Real regenerations happen at different wall-clock times, which XSSFWorkbook bakes into
        // core.xml's created/modified timestamps -- stamped explicitly here (rather than via
        // Thread.sleep) so the byte-level difference this test relies on isn't flaky under fast
        // back-to-back execution.
        var first = writeToBytes(ExampleSheetFixture.build(), Instant.parse("2024-01-01T00:00:00Z"));
        var second = writeToBytes(ExampleSheetFixture.build(), Instant.parse("2024-06-15T12:00:00Z"));

        assertFalse(Arrays.equals(first, second),
            "workbooks stamped with different creation times are expected to differ at the byte level "
                + "(this is the premise SpreadsheetAssert exists to work around)");

        assertEquals(first, second);
    }

    @Test
    void aDifferentCellValueIsReported() throws Exception {
        var expected = writeToBytes(ExampleSheetFixture.build());

        var actualWorkbook = ExampleSheetFixture.build();
        actualWorkbook.getSheetAt(0)
            .getRow(ExampleSheetFixture.ROW_DATA_1)
            .getCell(ExampleSheetFixture.COL_LABEL)
            .setCellValue("Item One (renamed)");
        var actual = writeToBytes(actualWorkbook);

        var failure = assertThrows(AssertionError.class, () -> assertEquals(expected, actual));
        assertTrue(failure.getMessage().contains("Item One"),
            "failure message should mention the differing text: " + failure.getMessage());
    }

    @Test
    void aCustomMessageIsPrefixedToTheFailure() throws Exception {
        var expected = writeToBytes(ExampleSheetFixture.build());

        var actualWorkbook = ExampleSheetFixture.build();
        actualWorkbook.getSheetAt(0)
            .getRow(ExampleSheetFixture.ROW_DATA_1)
            .getCell(ExampleSheetFixture.COL_LABEL)
            .setCellValue("Item One (renamed)");
        var actual = writeToBytes(actualWorkbook);

        var failure = assertThrows(AssertionError.class,
            () -> xlsxtopptx.testing.SpreadsheetAssert.assertEquals("budget export regression", expected, actual));
        assertTrue(failure.getMessage().startsWith("budget export regression"));
    }

    private static byte[] writeToBytes(org.apache.poi.xssf.usermodel.XSSFWorkbook workbook) throws Exception {
        try (workbook; var out = new ByteArrayOutputStream()) {
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] writeToBytes(org.apache.poi.xssf.usermodel.XSSFWorkbook workbook, Instant created) throws Exception {
        try (workbook; var out = new ByteArrayOutputStream()) {
            workbook.getProperties().getCoreProperties().setCreated(Optional.of(Date.from(created)));
            workbook.write(out);
            return out.toByteArray();
        }
    }
}
