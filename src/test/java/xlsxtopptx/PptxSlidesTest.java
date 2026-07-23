package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.awt.geom.Rectangle2D;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

class PptxSlidesTest {

    @Test
    void copiesShapesFromSource() throws Exception {
        try (var source = new XMLSlideShow(); var destination = new XMLSlideShow()) {
            var sourceSlide = source.createSlide();
            var box = sourceSlide.createTextBox();
            box.setAnchor(new Rectangle2D.Double(10, 10, 100, 20));
            box.addNewTextParagraph().addNewTextRun().setText("hello");

            var newSlide = PptxSlides.copyInto(destination, sourceSlide);

            var textBox = (XSLFTextBox) newSlide.getShapes().get(0);
            assertEquals("hello", textBox.getText());
        }
    }

    @Test
    void copiesBackgroundResolvedFromTheSourceMaster() throws Exception {
        // Mirrors how template.pptx itself is built: XSLFSlide.getBackground() falls back to the
        // slide's master when the slide has no background of its own, so copyInto must resolve
        // through that fallback rather than only checking the slide directly.
        try (var source = new XMLSlideShow(); var destination = new XMLSlideShow()) {
            var sourceSlide = source.createSlide();
            sourceSlide.getMasterSheet().getBackground().setFillColor(Color.BLUE);

            var newSlide = PptxSlides.copyInto(destination, sourceSlide);

            assertEquals(Color.BLUE, newSlide.getBackground().getFillColor());
        }
    }

    @Test
    void eachCopiedSlideKeepsItsOwnDistinctBackgroundWhenSharingOneMaster() throws Exception {
        // Regression test: destination.createSlide() reuses one shared master/layout for every
        // slide it creates. Applying a resolved background color via the plain getBackground()
        // fallback path (instead of forcing it onto the copied slide itself) would mutate that
        // shared master, so the second copy's color would silently overwrite the first copy's.
        try (var sourceA = new XMLSlideShow();
             var sourceB = new XMLSlideShow();
             var destination = new XMLSlideShow()) {

            var slideA = sourceA.createSlide();
            slideA.getBackground().setFillColor(Color.RED);
            var slideB = sourceB.createSlide();
            slideB.getBackground().setFillColor(Color.BLUE);

            var newSlideA = PptxSlides.copyInto(destination, slideA);
            var newSlideB = PptxSlides.copyInto(destination, slideB);

            assertEquals(Color.RED, newSlideA.getBackground().getFillColor());
            assertEquals(Color.BLUE, newSlideB.getBackground().getFillColor());
            assertNotEquals(newSlideA.getBackground().getFillColor(), newSlideB.getBackground().getFillColor());
        }
    }

    @Test
    void copiesUntouchedDefaultBackgroundWithoutError() throws Exception {
        // A slide that never had its background touched still resolves to a concrete color (the
        // theme's default background, white) rather than null -- copyInto must handle that case
        // too, not just explicitly-colored backgrounds.
        try (var source = new XMLSlideShow(); var destination = new XMLSlideShow()) {
            var sourceSlide = source.createSlide();

            var newSlide = PptxSlides.copyInto(destination, sourceSlide);

            assertEquals(sourceSlide.getBackground().getFillColor(), newSlide.getBackground().getFillColor());
        }
    }
}
