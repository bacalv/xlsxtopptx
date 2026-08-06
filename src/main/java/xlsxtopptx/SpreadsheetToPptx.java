package xlsxtopptx;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.geom.Rectangle2D;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Converts one Excel sheet (.xlsx) into a SINGLE PowerPoint slide, scaling everything (font size,
 * column widths, row heights) uniformly so the whole sheet fits on one slide -- no pagination.
 * Preserves cell fill colors, borders, font styling, alignment, and merged cells (including
 * vertical merges that span many rows).
 *
 * <p>The work is split across a few small, focused classes:
 * <ul>
 *   <li>{@link SheetReader} -- reads a POI {@link Sheet} into a POI-independent
 *       {@link SheetSnapshot}.
 *   <li>{@link LayoutEngine} -- computes the uniform font scale and resulting column/row sizes.
 *   <li>{@link SlideRenderer} -- writes a sized snapshot onto a POI slide/table.
 *   <li>{@link MergeIndex}, {@link MergeRegion}, {@link CellPos} -- merge-region bookkeeping.
 *   <li>{@link CellSnapshot}, {@link RowSnapshot}, {@link CellBorders}, {@link BorderEdgeStyle}
 *       -- the immutable per-cell/per-row data model.
 *   <li>{@link LayoutConstants} -- the tunable constants shared by all of the above.
 * </ul>
 *
 * Limitations (kept simple on purpose):
 * <ul>
 *   <li>.xlsx only (uses XSSFWorkbook).
 *   <li>A merged region straddling into columns beyond the sheet's detected column count, or
 *       otherwise malformed, is skipped defensively.
 *   <li>Theme-based colors fall back to a default (resolving a theme color to RGB needs the
 *       workbook's theme part, not just the cell style).
 *   <li>Very large sheets will end up with very small text -- that's the explicit trade-off of
 *       "everything on one slide".
 * </ul>
 */
@Slf4j
public final class SpreadsheetToPptx {

    public byte[] convert(ConvertParams params) throws IOException {
        try (var workbook = new XSSFWorkbook(params.getExcelInput());
             var ppt = openPresentation(params.getTemplateInput())) {

            var sheet = workbook.getSheetAt(params.getSheetIndex());
            var bounds = resolveBounds(sheet, params.getTargetRange());

            var slide = resolveSlide(ppt, params.getTemplateInput());
            if (params.getTitle() != null && !params.getTitle().isBlank()) {
                SlideRenderer.addTitle(ppt, slide, params.getTitle());
            }
            if (params.getLinkUrl() != null && !params.getLinkUrl().isBlank()) {
                SlideRenderer.addSourceLink(ppt, slide, params.getLinkUrl(), params.getLinkText());
            }

            var area = resolveTableArea(ppt, params.getTableArea());
            layoutTable(slide, sheet, bounds, area, params.isScaleToFitWidth(), params.isScaleToFitHeight());

            var out = new ByteArrayOutputStream();
            ppt.write(out);
            return out.toByteArray();
        }
    }

    /** Renders {@code bounds} of {@code sheet} as a table on {@code slide}, positioned and sized
     *  to exactly fill {@code area} (points, top-left origin) -- or does nothing if that range is
     *  empty. A sheet narrower than {@code area} with {@code scaleToFitWidth} off (i.e. the content
     *  is naturally narrow) is centered horizontally within {@code area} instead of hugging its
     *  left edge. */
    static void layoutTable(XSLFSlide slide, Sheet sheet, CellRangeAddress bounds, Rectangle2D area,
                             boolean scaleToFitWidth, boolean scaleToFitHeight) {
        var sheetSnapshot = SheetReader.read(sheet,
            bounds.getFirstRow(), bounds.getLastRow(), bounds.getFirstColumn(), bounds.getLastColumn());
        if (sheetSnapshot.rows().isEmpty()) return;

        var mergeIndex = MergeIndex.build(sheetSnapshot.merges());
        var layout = LayoutEngine.compute(sheetSnapshot, mergeIndex, area.getWidth(), area.getHeight(),
            scaleToFitWidth, scaleToFitHeight);

        var table = slide.createTable(sheetSnapshot.rowCount(), sheetSnapshot.numCols());
        var tableX = area.getX() + Math.max(0, (area.getWidth() - layout.tableWidth()) / 2);
        table.setAnchor(new Rectangle2D.Double(tableX, area.getY(), layout.tableWidth(), layout.tableHeight()));

        SlideRenderer.renderTable(table, sheetSnapshot, mergeIndex, layout.fontScale());

        for (int c = 0; c < sheetSnapshot.numCols(); c++) {
            table.setColumnWidth(c, layout.colWidths().get(c));
        }
        for (int r = 0; r < sheetSnapshot.rowCount(); r++) {
            table.getRows().get(r).setHeight(Math.max(layout.rowHeights().get(r), LayoutConstants.MIN_ROW_HEIGHT_PT));
        }

        for (var merge : sheetSnapshot.merges()) {
            if (merge.spansMultipleCells()) {
                table.mergeCells(merge.firstRow(), merge.lastRow(), merge.firstCol(), merge.lastCol());
            }
        }

        log.info("sheet={} rows={} cols={} merges={} fontScale={}",
            sheet.getSheetName(), sheetSnapshot.rowCount(), sheetSnapshot.numCols(),
            sheetSnapshot.merges().size(), layout.fontScale());
    }

    /** Resolves the inclusive, zero-based row/column bounds to convert: the parsed
     *  {@code targetRange} (A1 notation, e.g. "B2:F10") if given, otherwise the sheet's full used
     *  area. */
    private CellRangeAddress resolveBounds(Sheet sheet, String targetRange) {
        if (targetRange != null && !targetRange.isBlank()) {
            return CellRangeAddress.valueOf(targetRange);
        }
        return SheetReader.usedBounds(sheet);
    }

    /** A blank presentation, or one based on {@code templateInput} if given. */
    private XMLSlideShow openPresentation(InputStream templateInput) throws IOException {
        return templateInput != null ? new XMLSlideShow(templateInput) : new XMLSlideShow();
    }

    /** A blank new slide, or the template's first (and expected only) slide if a template was used. */
    private XSLFSlide resolveSlide(XMLSlideShow ppt, InputStream templateInput) {
        return templateInput != null ? ppt.getSlides().get(0) : ppt.createSlide();
    }

    /** The rectangle (in points) the table is laid out into: the caller's explicit
     *  {@code tableArea} if given, otherwise the full page inside the standard margins, minus
     *  room for a title. */
    private Rectangle2D resolveTableArea(XMLSlideShow ppt, Rectangle2D tableArea) {
        if (tableArea != null) return tableArea;

        var pageSize = ppt.getPageSize();
        return new Rectangle2D.Double(
            LayoutConstants.MARGIN_PT, LayoutConstants.MARGIN_PT + LayoutConstants.TITLE_HEIGHT_PT,
            pageSize.getWidth() - 2 * LayoutConstants.MARGIN_PT - LayoutConstants.OVERFLOW_SAFETY_PT,
            pageSize.getHeight() - 2 * LayoutConstants.MARGIN_PT - LayoutConstants.TITLE_HEIGHT_PT - LayoutConstants.OVERFLOW_SAFETY_PT);
    }
}
