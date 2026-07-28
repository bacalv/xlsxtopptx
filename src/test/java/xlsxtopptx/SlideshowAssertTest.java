package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.Instant;
import java.util.Arrays;
import java.util.Date;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static xlsxtopptx.testing.SlideshowAssert.assertEquals;

/**
 * Exercises {@link xlsxtopptx.testing.SlideshowAssert} (in the testFixtures source set) against
 * the real {@link SpreadsheetToPptx#convert} pipeline: running the same conversion twice yields
 * different bytes, but should still compare equal, while an actual content change should still be
 * caught.
 */
class SlideshowAssertTest {

    @Test
    void regeneratingTheSamePresentationProducesDifferentBytesButComparesEqual() throws Exception {
        var xlsxBytes = writeToBytes(ExampleSheetFixture.build());
        var base = convert(xlsxBytes, b -> b);

        // Real regenerations happen at different wall-clock times, which XMLSlideShow bakes into
        // core.xml's created/modified timestamps -- stamped explicitly here (rather than via
        // Thread.sleep) so the byte-level difference this test relies on isn't flaky under fast
        // back-to-back execution.
        var first = withCreatedTimestamp(base, Instant.parse("2024-01-01T00:00:00Z"));
        var second = withCreatedTimestamp(base, Instant.parse("2024-06-15T12:00:00Z"));

        assertFalse(Arrays.equals(first, second),
            "presentations stamped with different creation times are expected to differ at the byte level "
                + "(this is the premise SlideshowAssert exists to work around)");

        assertEquals(first, second);
    }

    @Test
    void aDifferentTitleIsReported() throws Exception {
        var xlsxBytes = writeToBytes(ExampleSheetFixture.build());

        var expected = convert(xlsxBytes, b -> b.title("Budget Offsite"));
        var actual = convert(xlsxBytes, b -> b.title("Budget Offsite (v2)"));

        var failure = assertThrows(AssertionError.class, () -> assertEquals(expected, actual));
        assertTrue(failure.getMessage().contains("Budget Offsite"),
            "failure message should mention the differing title: " + failure.getMessage());
    }

    @Test
    void aCustomMessageIsPrefixedToTheFailure() throws Exception {
        var xlsxBytes = writeToBytes(ExampleSheetFixture.build());

        var expected = convert(xlsxBytes, b -> b.title("Budget Offsite"));
        var actual = convert(xlsxBytes, b -> b.title("Budget Offsite (v2)"));

        var failure = assertThrows(AssertionError.class,
            () -> xlsxtopptx.testing.SlideshowAssert.assertEquals("slide deck regression", expected, actual));
        assertTrue(failure.getMessage().startsWith("slide deck regression"));
    }

    private static byte[] convert(byte[] xlsxBytes,
                                   java.util.function.UnaryOperator<ConvertParams.ConvertParamsBuilder> customizer) throws Exception {
        try (var excelInput = new ByteArrayInputStream(xlsxBytes)) {
            var params = customizer.apply(ConvertParams.builder().excelInput(excelInput).sheetIndex(0)).build();
            return new SpreadsheetToPptx().convert(params);
        }
    }

    private static byte[] writeToBytes(org.apache.poi.xssf.usermodel.XSSFWorkbook workbook) throws Exception {
        try (workbook; var out = new ByteArrayOutputStream()) {
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] withCreatedTimestamp(byte[] pptxBytes, Instant created) throws Exception {
        try (var in = new ByteArrayInputStream(pptxBytes);
             var ppt = new XMLSlideShow(in);
             var out = new ByteArrayOutputStream()) {
            ppt.getProperties().getCoreProperties().setCreated(Optional.of(Date.from(created)));
            ppt.write(out);
            return out.toByteArray();
        }
    }
}
