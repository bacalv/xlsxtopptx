package xlsxtopptx;

import lombok.Builder;
import lombok.Getter;

import java.awt.geom.Rectangle2D;
import java.io.InputStream;

/**
 * Parameters for {@link SpreadsheetToPptx#convert}.
 *
 * <p>{@code excelInput} is fully read by {@code convert}; opening it is the caller's
 * responsibility, but closing it afterward is not -- POI's {@code XSSFWorkbook} reads the stream
 * to completion and closes it itself as part of construction, so the stream is unusable (and
 * safe to close again, as a no-op) by the time {@code convert} returns.
 *
 * <p>{@code targetRange}, when set, is an Excel A1-notation range (e.g. {@code "B2:F10"})
 * restricting the conversion to that rectangle of the sheet instead of its full used area.
 *
 * <p>{@code scaleToFitWidth}/{@code scaleToFitHeight} opt into stretching column widths / row
 * heights to fill the slide's available width/height. Left off, a sheet narrower or shorter than
 * the page keeps its natural size instead of being stretched -- a narrower-than-page table is
 * then centered horizontally.
 *
 * <p>{@code templateInput}, when set, is a .pptx whose first slide is used as the base for the
 * output (e.g. one carrying a logo or other branding) instead of a blank slide on a blank
 * presentation. Like {@code excelInput}, it's fully read and left closed by {@code convert}. Left
 * unset, a blank presentation and slide are created as before.
 *
 * <p>{@code tableArea}, when set, is the exact rectangle (in points, relative to the slide's
 * top-left corner) the table is laid out into, overriding the default margin-based area. Left
 * unset, the default area is used -- the full page inside the standard margins, minus room for a
 * title if one is set.
 *
 * <p>{@code title}, when set, is rendered as a bold title textbox at the top of the slide. Left
 * unset (the default), no title is rendered at all.
 *
 * <p>{@code linkUrl}, when set, is rendered as a small hyperlinked textbox in the slide's bottom
 * margin -- e.g. a link to wherever the source spreadsheet can be downloaded -- with {@code
 * linkText} as its visible label, or {@code linkUrl} itself if {@code linkText} is left unset.
 * Left unset (the default), no link is rendered at all.
 */
@Getter
@Builder
public class ConvertParams {
    private final InputStream excelInput;
    @Builder.Default
    private final int sheetIndex = 0;
    @Builder.Default
    private final boolean scaleToFitWidth = false;
    @Builder.Default
    private final boolean scaleToFitHeight = false;
    private final String targetRange;
    private final InputStream templateInput;
    private final Rectangle2D tableArea;
    private final String title;
    private final String linkUrl;
    private final String linkText;
}
