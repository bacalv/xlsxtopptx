package xlsxtopptx;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class SlideRendererTest {

    private static CellSnapshot cell(String text, boolean bold, CellBorders borders) {
        return new CellSnapshot(text, null, LayoutConstants.DEFAULT_FONT_COLOR, bold, false,
            11.0, "Calibri", TextAlign.LEFT, borders);
    }

    @Test
    void titleTextAndFormattingAreSetCorrectly() throws Exception {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            SlideRenderer.addTitle(ppt, slide, "Funding Requests");

            var textBox = (XSLFTextBox) slide.getShapes().stream()
                .filter(XSLFTextBox.class::isInstance)
                .findFirst()
                .orElseThrow();

            assertEquals("Funding Requests", textBox.getText());
            var para = textBox.getTextParagraphs().get(0);
            var run = para.getTextRuns().get(0);
            assertEquals(16.0, run.getFontSize());
            assertTrue(run.isBold());
        }
    }

    @Test
    void titleParagraphDoesNotInheritMasterSpacingDefaults() throws Exception {
        // The slide master's default body style bakes in spcBef ~20% of font size, meant for
        // bulleted body text. Left unset, a title paragraph would inherit it and look
        // "double-spaced" -- addTitle must pin these to neutral explicitly.
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            SlideRenderer.addTitle(ppt, slide, "Title");

            var textBox = (XSLFTextBox) slide.getShapes().stream()
                .filter(XSLFTextBox.class::isInstance)
                .findFirst()
                .orElseThrow();
            var para = textBox.getTextParagraphs().get(0);

            assertEquals(0.0, para.getSpaceBefore());
            assertEquals(0.0, para.getSpaceAfter());
            assertEquals(100.0, para.getLineSpacing());
        }
    }

    @Test
    void titleHasNoDuplicateLeadingParagraphOrRun() throws Exception {
        // XSLFTextBox.createTextBox() pre-populates a default empty paragraph/run; addTitle must
        // reuse it rather than appending a second one alongside it.
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            SlideRenderer.addTitle(ppt, slide, "Title");

            var textBox = (XSLFTextBox) slide.getShapes().stream()
                .filter(XSLFTextBox.class::isInstance)
                .findFirst()
                .orElseThrow();

            assertEquals(1, textBox.getTextParagraphs().size());
            assertEquals(1, textBox.getTextParagraphs().get(0).getTextRuns().size());
        }
    }

    @Test
    void cellTextAndFontScaleAreApplied() throws Exception {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            var table = slide.createTable(1, 1);
            var sheet = new SheetSnapshot(
                List.of(new RowSnapshot(List.of(cell("2,145.00", true, CellBorders.NONE)))),
                List.of(), 1);

            SlideRenderer.renderTable(table, sheet, MergeIndex.build(List.of()), 0.5);

            var textCell = table.getCell(0, 0);
            assertEquals("2,145.00", textCell.getText());
            var run = textCell.getTextParagraphs().get(0).getTextRuns().get(0);
            assertEquals(5.5, run.getFontSize()); // 11pt base * 0.5 fontScale
            assertTrue(run.isBold());
        }
    }

    @Test
    void endParaRPrIsScaledToMatchTheVisibleRunSize() throws Exception {
        // Regression test: a paragraph's trailing endParaRPr defaults to POI's stock 11pt. Viewers
        // use the larger of the run size and endParaRPr when computing line height, so leaving it
        // unscaled made every row render at ~11pt tall regardless of the actual (scaled-down) font.
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            var table = slide.createTable(1, 1);
            var sheet = new SheetSnapshot(
                List.of(new RowSnapshot(List.of(cell("x", false, CellBorders.NONE)))),
                List.of(), 1);

            SlideRenderer.renderTable(table, sheet, MergeIndex.build(List.of()), 0.5);

            var para = table.getCell(0, 0).getTextParagraphs().get(0);
            var endParaRPrSz = para.getXmlObject().getEndParaRPr().getSz();
            assertEquals(550, endParaRPrSz); // 5.5pt * 100 (hundredths of a point)
        }
    }

    @Test
    void cellParagraphDoesNotInheritMasterSpacingDefaults() throws Exception {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            var table = slide.createTable(1, 1);
            var sheet = new SheetSnapshot(
                List.of(new RowSnapshot(List.of(cell("x", false, CellBorders.NONE)))),
                List.of(), 1);

            SlideRenderer.renderTable(table, sheet, MergeIndex.build(List.of()), 1.0);

            var para = table.getCell(0, 0).getTextParagraphs().get(0);
            assertEquals(0.0, para.getSpaceBefore());
            assertEquals(0.0, para.getSpaceAfter());
            assertEquals(100.0, para.getLineSpacing());
        }
    }

    @Test
    void visibleBorderEdgeIsWrittenWithItsStyleAndColor() throws Exception {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            var table = slide.createTable(1, 1);
            var lightBlue = new Color(0xAD, 0xD8, 0xE6);
            var borders = new CellBorders(
                BorderEdgeStyle.NONE,
                new BorderEdgeStyle(org.apache.poi.ss.usermodel.BorderStyle.THIN, lightBlue),
                BorderEdgeStyle.NONE,
                BorderEdgeStyle.NONE
            );
            var sheet = new SheetSnapshot(
                List.of(new RowSnapshot(List.of(cell("x", false, borders)))),
                List.of(), 1);

            SlideRenderer.renderTable(table, sheet, MergeIndex.build(List.of()), 1.0);

            var renderedCell = table.getCell(0, 0);
            assertEquals(lightBlue, renderedCell.getBorderColor(BorderEdge.bottom));
            assertEquals(0.75, renderedCell.getBorderWidth(BorderEdge.bottom));
        }
    }

    @Test
    void coveredCellsAreLeftBlankInsteadOfBeingStyled() throws Exception {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            var table = slide.createTable(1, 2);
            var sheet = new SheetSnapshot(
                List.of(new RowSnapshot(List.of(cell("Banner", false, CellBorders.NONE), cell("", false, CellBorders.NONE)))),
                List.of(new MergeRegion(0, 0, 0, 1)), 2);
            var mergeIndex = MergeIndex.build(sheet.merges());

            SlideRenderer.renderTable(table, sheet, mergeIndex, 1.0);

            assertEquals("Banner", table.getCell(0, 0).getText());
            assertTrue(table.getCell(0, 1).getTextParagraphs().isEmpty());
        }
    }
}
