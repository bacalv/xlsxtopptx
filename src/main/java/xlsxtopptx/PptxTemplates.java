package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/** Loads a single slide out of a multi-slide .pptx classpath resource as a standalone
 *  single-slide presentation -- ready to use as {@link ConvertParams#getTemplateInput()}, which
 *  always uses the template's first (and only) slide. */
public final class PptxTemplates {
    private PptxTemplates() {}

    /** {@code slideIndex} is 0-based, e.g. {@code loadFromResource("/pptx/template.pptx", 2)}
     *  loads the third slide. */
    public static byte[] loadFromResource(String resourcePath, int slideIndex) throws IOException {
        try (var in = PptxTemplates.class.getResourceAsStream(resourcePath)) {
            if (in == null) {
                throw new IOException("Resource not found on classpath: " + resourcePath);
            }
            try (var source = new XMLSlideShow(in);
                 var extracted = new XMLSlideShow()) {

                extracted.setPageSize(source.getPageSize());
                PptxSlides.copyInto(extracted, source.getSlides().get(slideIndex));

                var out = new ByteArrayOutputStream();
                extracted.write(out);
                return out.toByteArray();
            }
        }
    }
}
