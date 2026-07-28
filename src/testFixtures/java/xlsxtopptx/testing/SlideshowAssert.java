package xlsxtopptx.testing;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.UncheckedIOException;

/**
 * Asserts that two .pptx slideshows are semantically equal: same slides, shapes (position/size),
 * text and run styling, table cell contents, and picture bytes -- while ignoring incidental
 * byte-level differences between two writer runs, such as core.xml timestamps, relationship id
 * numbering, or XML attribute ordering. Regenerating "the same" presentation from the same inputs
 * at a different time will produce different bytes but should still compare equal here.
 *
 * <p>Framework-agnostic: throws a plain {@link AssertionError} on failure, so it works the same
 * under JUnit, TestNG, or any other runner.
 */
public final class SlideshowAssert {
    private SlideshowAssert() {}

    /** @throws AssertionError if the two slideshows are not semantically equal. */
    public static void assertEquals(byte[] expectedPptx, byte[] actualPptx) {
        assertEquals(null, expectedPptx, actualPptx);
    }

    /** @throws AssertionError if the two slideshows are not semantically equal. {@code message} is
     *  prefixed to the failure, same as JUnit's message-first assertion overloads. */
    public static void assertEquals(String message, byte[] expectedPptx, byte[] actualPptx) {
        try (var expected = new XMLSlideShow(new ByteArrayInputStream(expectedPptx));
             var actual = new XMLSlideShow(new ByteArrayInputStream(actualPptx))) {
            var diffs = SlideshowComparator.compare(expected, actual);
            DiffReport.failIfAny(message, "slideshows are not semantically equal", diffs);
        } catch (IOException e) {
            throw new UncheckedIOException("failed to read slideshow bytes for comparison", e);
        }
    }
}
