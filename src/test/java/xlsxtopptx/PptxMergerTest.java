package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class PptxMergerTest {

    @Test
    void throwsOnEmptyInputList() {
        List<InputStream> emptyInputs = List.of();
        assertThrows(IllegalArgumentException.class, () -> PptxMerger.merge(emptyInputs));
    }

    @Test
    void concatenatesSlidesFromEachInputInOrder() throws Exception {
        var inputA = singleSlidePptx("A");
        var inputB = singleSlidePptx("B");

        var merged = mergeAndReload(List.of(new ByteArrayInputStream(inputA), new ByteArrayInputStream(inputB)));

        var slides = merged.getSlides();
        assertEquals(2, slides.size());
        assertEquals("A", textOf(slides.get(0)));
        assertEquals("B", textOf(slides.get(1)));
    }

    @Test
    void preservesSlideOrderAcrossMultiSlideInputs() throws Exception {
        var inputA = twoSlidePptx("A1", "A2");
        var inputB = singleSlidePptx("B1");

        var merged = mergeAndReload(List.of(new ByteArrayInputStream(inputA), new ByteArrayInputStream(inputB)));

        var slides = merged.getSlides();
        assertEquals(3, slides.size());
        assertEquals("A1", textOf(slides.get(0)));
        assertEquals("A2", textOf(slides.get(1)));
        assertEquals("B1", textOf(slides.get(2)));
    }

    @Test
    void pageSizeComesFromTheFirstInput() throws Exception {
        var widescreen = new Dimension(960, 540);
        var standard = new Dimension(720, 540);
        var inputA = singleSlidePptxWithPageSize(widescreen);
        var inputB = singleSlidePptxWithPageSize(standard);

        var merged = mergeAndReload(List.of(new ByteArrayInputStream(inputA), new ByteArrayInputStream(inputB)));

        assertEquals(widescreen, merged.getPageSize());
    }

    @Test
    void eachMergedSlideKeepsItsOwnBackground() throws Exception {
        var inputA = singleSlidePptxWithBackground(Color.RED);
        var inputB = singleSlidePptxWithBackground(Color.BLUE);

        var merged = mergeAndReload(List.of(new ByteArrayInputStream(inputA), new ByteArrayInputStream(inputB)));

        var slides = merged.getSlides();
        assertEquals(Color.RED, slides.get(0).getBackground().getFillColor());
        assertEquals(Color.BLUE, slides.get(1).getBackground().getFillColor());
    }

    @Test
    void leavesInputStreamsClosedAfterUse() throws Exception {
        // POI's XMLSlideShow reads a stream to completion and closes it itself as part of
        // construction (verified directly against the library rather than assumed) -- merge
        // doesn't need to, and can't avoid, closing each input itself.
        var closeTracking = new CloseTrackingInputStream(new ByteArrayInputStream(singleSlidePptx("A")));

        PptxMerger.merge(List.of(closeTracking));

        assertTrue(closeTracking.closed);
    }

    private static XMLSlideShow mergeAndReload(List<InputStream> inputs) throws IOException {
        var merged = PptxMerger.merge(inputs);
        return new XMLSlideShow(new ByteArrayInputStream(merged));
    }

    private static String textOf(org.apache.poi.xslf.usermodel.XSLFSlide slide) {
        return ((XSLFTextBox) slide.getShapes().get(0)).getText();
    }

    private static byte[] singleSlidePptx(String text) throws IOException {
        try (var ppt = new XMLSlideShow()) {
            addTextSlide(ppt, text);
            return toBytes(ppt);
        }
    }

    private static byte[] twoSlidePptx(String text1, String text2) throws IOException {
        try (var ppt = new XMLSlideShow()) {
            addTextSlide(ppt, text1);
            addTextSlide(ppt, text2);
            return toBytes(ppt);
        }
    }

    private static byte[] singleSlidePptxWithPageSize(Dimension pageSize) throws IOException {
        try (var ppt = new XMLSlideShow()) {
            ppt.setPageSize(pageSize);
            addTextSlide(ppt, "text");
            return toBytes(ppt);
        }
    }

    private static byte[] singleSlidePptxWithBackground(Color color) throws IOException {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            slide.getBackground().setFillColor(color);
            return toBytes(ppt);
        }
    }

    private static void addTextSlide(XMLSlideShow ppt, String text) {
        var slide = ppt.createSlide();
        var box = slide.createTextBox();
        box.setAnchor(new Rectangle2D.Double(10, 10, 300, 40));
        box.addNewTextParagraph().addNewTextRun().setText(text);
    }

    private static byte[] toBytes(XMLSlideShow ppt) throws IOException {
        var out = new ByteArrayOutputStream();
        ppt.write(out);
        return out.toByteArray();
    }

    private static final class CloseTrackingInputStream extends FilterInputStream {
        boolean closed = false;

        CloseTrackingInputStream(InputStream in) {
            super(in);
        }

        @Override
        public void close() throws IOException {
            closed = true;
            super.close();
        }
    }
}
