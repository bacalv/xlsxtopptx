package xlsxtopptx;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;

import java.awt.geom.Rectangle2D;

/** Writes a {@link SheetSnapshot}, already sized by {@link LayoutEngine}, onto a POI slide/table. */
public final class SlideRenderer {
    private SlideRenderer() {}

    public static void addTitle(XMLSlideShow ppt, XSLFSlide slide, String text) {
        var pageSize = ppt.getPageSize();
        var title = slide.createTextBox();
        title.setAnchor(new Rectangle2D.Double(
            LayoutConstants.MARGIN_PT, LayoutConstants.MARGIN_PT * 0.4,
            pageSize.getWidth() - 2 * LayoutConstants.MARGIN_PT, LayoutConstants.TITLE_HEIGHT_PT));

        var existingParagraphs = title.getTextParagraphs();
        var para = existingParagraphs.isEmpty() ? title.addNewTextParagraph() : existingParagraphs.get(0);
        var existingRuns = para.getTextRuns();
        var run = existingRuns.isEmpty() ? para.addNewTextRun() : existingRuns.get(0);

        para.setSpaceBefore(0.0);
        para.setSpaceAfter(0.0);
        para.setLineSpacing(100.0);
        run.setText(text);
        run.setFontSize(16.0);
        run.setBold(true);
        setEndParaRPrSize(para, 16.0);
    }

    public static void renderTable(XSLFTable table, SheetSnapshot sheet, MergeIndex mergeIndex, double fontScale) {
        var rows = sheet.rows();
        for (int r = 0; r < rows.size(); r++) {
            writeRow(table, r, rows.get(r), mergeIndex, fontScale);
        }
    }

    private static void writeRow(XSLFTable table, int rowIdx, RowSnapshot rowSnap, MergeIndex mergeIndex, double fontScale) {
        var tr = table.getRows().get(rowIdx);
        for (int c = 0; c < rowSnap.cells().size(); c++) {
            if (mergeIndex.isCovered(new CellPos(rowIdx, c))) continue; // hidden once merged; leave blank
            var cell = tr.getCells().get(c);
            styleCell(cell, rowSnap.cells().get(c), fontScale);
        }
    }

    private static void styleCell(XSLFTableCell cell, CellSnapshot cs, double fontScale) {
        if (cs.fill() != null) {
            cell.setFillColor(cs.fill());
        }

        applyBorder(cell, BorderEdge.top, cs.borders().top());
        applyBorder(cell, BorderEdge.bottom, cs.borders().bottom());
        applyBorder(cell, BorderEdge.left, cs.borders().left());
        applyBorder(cell, BorderEdge.right, cs.borders().right());

        // Pin insets explicitly instead of leaving them unset: different viewers (PowerPoint,
        // Keynote, Google Slides) apply their own default cell padding when it's left unspecified,
        // and those defaults don't agree with each other or with the padding this class's own
        // width/height math assumes -- that mismatch is what caused text to wrap and rows to
        // overflow the slide in some viewers but not others.
        cell.setLeftInset(LayoutConstants.CELL_PADDING_PT / 2);
        cell.setRightInset(LayoutConstants.CELL_PADDING_PT / 2);
        cell.setTopInset(LayoutConstants.ROW_PADDING_PT / 2);
        cell.setBottomInset(LayoutConstants.ROW_PADDING_PT / 2);

        var existing = cell.getTextParagraphs();
        var para = existing.isEmpty() ? cell.addNewTextParagraph() : existing.get(0);
        para.setTextAlign(cs.align());
        // The slide master's default body paragraph style bakes in spcBef (space before a
        // paragraph, ~20% of font size) meant for bulleted body text. Table cell paragraphs
        // inherit it unless overridden, stacking extra vertical space on top of the line height
        // this class already accounts for -- the "double-spaced" look. Pin all three to neutral.
        para.setSpaceBefore(0.0);
        para.setSpaceAfter(0.0);
        para.setLineSpacing(100.0);

        var run = para.addNewTextRun();
        run.setText(cs.text() == null ? "" : cs.text());
        var baseFont = cs.fontSize() > 0 ? cs.fontSize() : LayoutConstants.DEFAULT_FONT_PT;
        var scaledFont = Math.max(LayoutConstants.MIN_FONT_PT, baseFont * fontScale);
        run.setFontSize(scaledFont);
        run.setFontFamily(cs.fontName() != null ? cs.fontName() : LayoutConstants.DEFAULT_FONT_FAMILY);
        run.setBold(cs.bold());
        run.setItalic(cs.italic());
        run.setFontColor(cs.fontColor() != null ? cs.fontColor() : LayoutConstants.DEFAULT_FONT_COLOR);
        setEndParaRPrSize(para, scaledFont);
    }

    /**
     * A paragraph's trailing endParaRPr (the formatting of the "phantom" character at the
     * paragraph mark) defaults to POI's stock 11pt whenever left unset. PowerPoint and Keynote
     * both use the larger of a paragraph's run sizes and its endParaRPr size when computing the
     * line's rendered height -- so a table row whose visible text is scaled down to a few points
     * still lays out at ~11pt tall unless endParaRPr is scaled down to match, which is what
     * caused rows to overflow the slide (the requested row height was honored, but the actual
     * rendered line height wasn't).
     */
    private static void setEndParaRPrSize(XSLFTextParagraph para, double fontSizePt) {
        var xmlPara = para.getXmlObject();
        var endPr = xmlPara.isSetEndParaRPr() ? xmlPara.getEndParaRPr() : xmlPara.addNewEndParaRPr();
        endPr.setSz((int) Math.round(fontSizePt * 100));
    }

    private static void applyBorder(XSLFTableCell cell, BorderEdge edge, BorderEdgeStyle edgeStyle) {
        if (!edgeStyle.isVisible()) return;
        cell.setBorderWidth(edge, borderWidthPt(edgeStyle.style()));
        cell.setBorderColor(edge, edgeStyle.color() != null ? edgeStyle.color() : LayoutConstants.DEFAULT_BORDER_COLOR);
    }

    private static double borderWidthPt(BorderStyle style) {
        return switch (style) {
            case THICK -> 2.5;
            case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT -> 1.5;
            case DOUBLE -> 2.0;
            default -> 0.75;
        };
    }
}
