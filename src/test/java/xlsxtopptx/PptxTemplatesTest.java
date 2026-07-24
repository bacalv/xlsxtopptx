package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.awt.Dimension;
import java.io.ByteArrayInputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Exercises {@link PptxTemplates#loadFromResource} against
 * {@code src/test/resources/pptx/multi-slide-fixture.pptx}: three 960x540 slides, each with a
 * distinct label ("Slide Zero"/"Slide One"/"Slide Two") and background color (red/green/blue).
 */
class PptxTemplatesTest {
    private static final String FIXTURE = "/pptx/multi-slide-fixture.pptx";

    @Test
    void loadsTheSlideAtTheGivenIndex() throws Exception {
        var extracted = loadAndReload(FIXTURE, 1);

        var slide = extracted.getSlides().get(0);
        assertEquals("Slide One", textOf(slide));
        assertEquals(Color.GREEN, slide.getBackground().getFillColor());
    }

    @Test
    void loadsADifferentSlideForADifferentIndex() throws Exception {
        var first = loadAndReload(FIXTURE, 0);
        var last = loadAndReload(FIXTURE, 2);

        assertEquals("Slide Zero", textOf(first.getSlides().get(0)));
        assertEquals(Color.RED, first.getSlides().get(0).getBackground().getFillColor());

        assertEquals("Slide Two", textOf(last.getSlides().get(0)));
        assertEquals(Color.BLUE, last.getSlides().get(0).getBackground().getFillColor());
    }

    @Test
    void resultIsASingleSlidePresentationEvenThoughTheSourceHasThree() throws Exception {
        var extracted = loadAndReload(FIXTURE, 0);

        assertEquals(1, extracted.getSlides().size());
    }

    @Test
    void pageSizeMatchesTheSource() throws Exception {
        var extracted = loadAndReload(FIXTURE, 0);

        assertEquals(new Dimension(960, 540), extracted.getPageSize());
    }

    @Test
    void throwsWhenTheResourceIsMissingFromTheClasspath() {
        var ex = assertThrows(IOException.class, () -> PptxTemplates.loadFromResource("/pptx/does-not-exist.pptx", 0));
        assertEquals("Resource not found on classpath: /pptx/does-not-exist.pptx", ex.getMessage());
    }

    private static XMLSlideShow loadAndReload(String resourcePath, int slideIndex) throws IOException {
        var bytes = PptxTemplates.loadFromResource(resourcePath, slideIndex);
        return new XMLSlideShow(new ByteArrayInputStream(bytes));
    }

    private static String textOf(org.apache.poi.xslf.usermodel.XSLFSlide slide) {
        return ((XSLFTextBox) slide.getShapes().get(0)).getText();
    }
}
