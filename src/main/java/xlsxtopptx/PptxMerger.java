package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/** Concatenates the slides of several .pptx files, in the given order, into one presentation. */
public final class PptxMerger {
    private PptxMerger() {}

    /** Fully reads each stream; opening them is the caller's responsibility, but closing them
     *  afterward is not -- POI's {@code XMLSlideShow} reads a stream to completion and closes it
     *  itself as part of construction, so each stream is unusable (and safe to close again, as a
     *  no-op) by the time {@code merge} returns. Same behavior as {@link ConvertParams}. */
    public static byte[] merge(List<InputStream> pptxInputs) throws IOException {
        if (pptxInputs.isEmpty()) {
            throw new IllegalArgumentException("pptxInputs must not be empty");
        }

        try (var merged = new XMLSlideShow()) {
            var pageSizeSet = false;

            for (var in : pptxInputs) {
                try (var source = new XMLSlideShow(in)) {
                    if (!pageSizeSet) {
                        // PowerPoint requires one page size per file; slides from a source with a
                        // different aspect ratio are copied at their original coordinates, not
                        // rescaled, so mixing e.g. 4:3 and 16:9 inputs will misalign shapes.
                        merged.setPageSize(source.getPageSize());
                        pageSizeSet = true;
                    }
                    for (var sourceSlide : source.getSlides()) {
                        PptxSlides.copyInto(merged, sourceSlide);
                    }
                }
            }

            var out = new ByteArrayOutputStream();
            merged.write(out);
            return out.toByteArray();
        }
    }
}
