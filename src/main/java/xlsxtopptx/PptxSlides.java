package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/** Shared slide-copy logic used by {@link PptxMerger} and {@link PptxTemplates}. */
final class PptxSlides {
    private PptxSlides() {}

    /** Copies {@code source}'s shapes and resolved background onto a new slide in
     *  {@code destination}. */
    static XSLFSlide copyInto(XMLSlideShow destination, XSLFSlide source) {
        var newSlide = destination.createSlide();
        newSlide.importContent(source);

        // importContent only copies the shape tree -- the background may live on the slide
        // itself, its layout, or its master, so resolve it explicitly and re-apply it. Solid
        // fills only; gradient/picture backgrounds aren't carried over.
        var fillColor = source.getBackground().getFillColor();
        if (fillColor != null) {
            ownBackground(newSlide).setFillColor(fillColor);
        }
        return newSlide;
    }

    /** {@code slide.getBackground()}, but forced to live on the slide itself rather than
     *  falling back to its layout/master -- so that setting it doesn't leak onto every other
     *  slide sharing that same master. */
    private static XSLFBackground ownBackground(XSLFSlide slide) {
        var cSld = slide.getXmlObject().getCSld();
        if (!cSld.isSetBg()) {
            cSld.addNewBg();
        }
        return slide.getBackground();
    }
}
