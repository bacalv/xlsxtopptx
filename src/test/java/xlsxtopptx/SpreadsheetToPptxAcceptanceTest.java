package xlsxtopptx;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.UnaryOperator;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * End-to-end: generates the example sheet (see {@link ExampleSheetFixture}), runs the real
 * {@link SpreadsheetToPptx#convert} pipeline against it, then reopens the resulting .pptx and
 * checks it against everything discovered the hard way during manual testing this converter went
 * through -- merges, formatted numbers, the light-blue border regression, and the table fitting
 * on the slide without overflowing or being stretched to waste vertical space.
 */
class SpreadsheetToPptxAcceptanceTest {

    @TempDir
    Path tempDir;

    private Path xlsxPath;
    private Path pptxPath;
    private XMLSlideShow openedResult;

    @BeforeEach
    void generateFixtureWorkbook() throws Exception {
        xlsxPath = tempDir.resolve("example.xlsx");
        pptxPath = tempDir.resolve("example.pptx");

        try (var workbook = ExampleSheetFixture.build();
             var out = new FileOutputStream(xlsxPath.toFile())) {
            workbook.write(out);
        }
    }

    @AfterEach
    void closeOpenedResult() throws Exception {
        if (openedResult != null) openedResult.close();
    }

    @Test
    void convertsTheExampleSheetToASinglePowerPointSlide() throws Exception {
        convertFixture();

        assertTrue(Files.exists(pptxPath));

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            assertEquals(1, ppt.getSlides().size());
        }
    }

    @Test
    void noTitleIsRenderedByDefault() throws Exception {
        convertFixture();

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            var slide = ppt.getSlides().get(0);
            var hasTextBox = slide.getShapes().stream().anyMatch(XSLFTextBox.class::isInstance);

            assertFalse(hasTextBox, "no title textbox should be added unless title is explicitly set");
        }
    }

    @Test
    void titleIsRenderedWhenSupplied() throws Exception {
        convertFixture(b -> b.title("Custom Title"));

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            var slide = ppt.getSlides().get(0);
            var textBox = (XSLFTextBox) slide.getShapes().stream()
                .filter(XSLFTextBox.class::isInstance)
                .findFirst()
                .orElseThrow();

            assertEquals("Custom Title", textBox.getText());
        }
    }

    @Test
    void tableHasTheExpectedDimensions() throws Exception {
        convertFixture();

        var table = readBackTable();

        assertEquals(ExampleSheetFixture.ROW_COUNT, table.getNumberOfRows());
        assertEquals(ExampleSheetFixture.COL_COUNT, table.getNumberOfColumns());
    }

    @Test
    void mergedBannerCellsKeepTheirSpan() throws Exception {
        convertFixture();

        var table = readBackTable();

        var yearBanner = table.getCell(ExampleSheetFixture.ROW_YEAR_BANNER, 2);
        assertEquals(2, yearBanner.getGridSpan(), "FY 2026 Forecast banner should span 2 columns");
        assertTrue(table.getCell(ExampleSheetFixture.ROW_YEAR_BANNER, 3).isMerged(),
            "the cell to the right of the banner anchor should be hidden inside the merge");

        var requestBanner = table.getCell(ExampleSheetFixture.ROW_REQUEST_BANNER, 1);
        assertEquals(ExampleSheetFixture.COL_TOTAL, requestBanner.getGridSpan(),
            "the requests banner should span from column 1 through the Total column");
    }

    @Test
    void formattedNumbersAndLabelsSurviveTheRoundTrip() throws Exception {
        convertFixture();

        var table = readBackTable();

        assertEquals("Products & Functions Technology [L5]",
            table.getCell(ExampleSheetFixture.ROW_SUMMARY, ExampleSheetFixture.COL_LABEL).getText());
        assertEquals("2,145.00",
            table.getCell(ExampleSheetFixture.ROW_SUMMARY, ExampleSheetFixture.COL_FIRST_METRIC).getText());
        assertEquals("13,581.00",
            table.getCell(ExampleSheetFixture.ROW_SUMMARY, ExampleSheetFixture.COL_TOTAL).getText());

        assertEquals("Total Funding Requests",
            table.getCell(ExampleSheetFixture.ROW_TOTAL, ExampleSheetFixture.COL_LABEL).getText());
        assertEquals("104.00",
            table.getCell(ExampleSheetFixture.ROW_TOTAL, ExampleSheetFixture.COL_TOTAL).getText());

        // The deliberately blank spacer row must round-trip as blank, not "0" or garbage.
        assertEquals("",
            table.getCell(ExampleSheetFixture.ROW_SPACER, ExampleSheetFixture.COL_FIRST_METRIC).getText());
    }

    @Test
    void lightBlueBorderSurvivesEvenThoughApplyBorderWasUnset() throws Exception {
        // The fixture's summary row style explicitly unsets applyBorder to simulate real-world
        // writers like openpyxl -- this is the regression test for that fix, run end-to-end.
        convertFixture();

        var table = readBackTable();
        var lightBlue = new Color(0xAD, 0xD8, 0xE6);

        for (int col = ExampleSheetFixture.COL_FIRST_METRIC; col <= ExampleSheetFixture.COL_TOTAL; col++) {
            var cell = table.getCell(ExampleSheetFixture.ROW_SUMMARY, col);
            assertEquals(lightBlue, cell.getBorderColor(BorderEdge.bottom),
                "column " + col + " of the summary row should carry the light-blue bottom border");
        }
    }

    @Test
    void ordinaryDataRowsHaveNoBorder() throws Exception {
        convertFixture();

        var table = readBackTable();

        var cell = table.getCell(ExampleSheetFixture.ROW_DATA_1, ExampleSheetFixture.COL_FIRST_METRIC);
        assertNull(cell.getBorderColor(BorderEdge.bottom));
    }

    @Test
    void tableFitsWithinTheSlideBoundsWithoutBeingStretchedToFillIt() throws Exception {
        convertFixture();

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            var pageSize = ppt.getPageSize();
            var slide = ppt.getSlides().get(0);
            var table = (XSLFTable) slide.getShapes().stream()
                .filter(XSLFTable.class::isInstance)
                .findFirst()
                .orElseThrow();
            var anchor = table.getAnchor();

            assertTrue(anchor.getMaxX() <= pageSize.getWidth() + 0.5,
                "table must not overflow the right edge of the slide");
            assertTrue(anchor.getMaxY() <= pageSize.getHeight() + 0.5,
                "table must not overflow the bottom edge of the slide");

            // This eight-row fixture has ample room on a standard slide, so it should NOT be
            // stretched to consume the entire available height -- that was the "wasted space per
            // row" bug. Leave generous headroom in the assertion since the exact figure depends on
            // tunable layout constants, not just this behavior.
            var availableHeight = pageSize.getHeight() - 2 * LayoutConstants.MARGIN_PT - LayoutConstants.TITLE_HEIGHT_PT;
            assertTrue(anchor.getHeight() < availableHeight * 0.9,
                "a short fixture like this should end up noticeably shorter than the slide, not stretched to fill it");
        }
    }

    @Test
    void narrowTableIsCenteredHorizontallyWhenNotScaledToFitWidth() throws Exception {
        convertFixture();

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            var pageSize = ppt.getPageSize();
            var slide = ppt.getSlides().get(0);
            var table = (XSLFTable) slide.getShapes().stream()
                .filter(XSLFTable.class::isInstance)
                .findFirst()
                .orElseThrow();
            var anchor = table.getAnchor();

            var availableWidth = pageSize.getWidth() - 2 * LayoutConstants.MARGIN_PT - LayoutConstants.OVERFLOW_SAFETY_PT;
            assertTrue(anchor.getWidth() < availableWidth,
                "this fixture's natural width should end up narrower than the slide by default, leaving room to center");

            var expectedX = LayoutConstants.MARGIN_PT + (availableWidth - anchor.getWidth()) / 2;
            assertEquals(expectedX, anchor.getX(), 0.5,
                "a table narrower than the slide should be centered horizontally, not left-aligned");
        }
    }

    @Test
    void scaleToFitWidthStretchesTheTableToTheFullAvailableWidth() throws Exception {
        convertFixture(b -> b.scaleToFitWidth(true));

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            var pageSize = ppt.getPageSize();
            var slide = ppt.getSlides().get(0);
            var table = (XSLFTable) slide.getShapes().stream()
                .filter(XSLFTable.class::isInstance)
                .findFirst()
                .orElseThrow();
            var anchor = table.getAnchor();

            var availableWidth = pageSize.getWidth() - 2 * LayoutConstants.MARGIN_PT - LayoutConstants.OVERFLOW_SAFETY_PT;
            assertEquals(availableWidth, anchor.getWidth(), 0.5,
                "scaleToFitWidth should stretch the table to fill the available width");
            assertEquals(LayoutConstants.MARGIN_PT, anchor.getX(), 0.5,
                "a table filling the available width should sit at the left margin, not be centered");
        }
    }

    @Test
    void scaleToFitHeightStretchesTheTableToTheFullAvailableHeight() throws Exception {
        convertFixture(b -> b.scaleToFitHeight(true));

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            var pageSize = ppt.getPageSize();
            var slide = ppt.getSlides().get(0);
            var table = (XSLFTable) slide.getShapes().stream()
                .filter(XSLFTable.class::isInstance)
                .findFirst()
                .orElseThrow();
            var anchor = table.getAnchor();

            var availableHeight = pageSize.getHeight() - 2 * LayoutConstants.MARGIN_PT
                - LayoutConstants.TITLE_HEIGHT_PT - LayoutConstants.OVERFLOW_SAFETY_PT;
            assertEquals(availableHeight, anchor.getHeight(), 0.5,
                "scaleToFitHeight should stretch the table to fill the available height");
        }
    }

    @Test
    void targetRangeRestrictsConversionToTheRequestedRectangle() throws Exception {
        // A6:F7 in Excel's 1-based A1 notation is the fixture's two plain data rows (indices 5-6)
        // across all six columns -- everything else (banners, headers, summary, total) is excluded.
        convertFixture(b -> b.targetRange("A6:F7"));

        var table = readBackTable();

        assertEquals(2, table.getNumberOfRows());
        assertEquals(ExampleSheetFixture.COL_COUNT, table.getNumberOfColumns());
        assertEquals("Item One", table.getCell(0, ExampleSheetFixture.COL_LABEL).getText());
        assertEquals("Item Two", table.getCell(1, ExampleSheetFixture.COL_LABEL).getText());
    }

    @Test
    void tableAreaOverridesTheDefaultTablePosition() throws Exception {
        var customArea = new Rectangle2D.Double(50, 120, 300, 200);
        convertFixture(b -> b.tableArea(customArea));

        var table = readBackTable();
        var anchor = table.getAnchor();

        assertEquals(customArea.getY(), anchor.getY(), 0.5,
            "table should sit at the top of the custom tableArea, not the default margin position");
        assertTrue(anchor.getX() >= customArea.getX() - 0.5,
            "table should not start left of the custom tableArea");
        assertTrue(anchor.getMaxX() <= customArea.getMaxX() + 0.5,
            "table should not overflow the right edge of the custom tableArea");
        assertTrue(anchor.getHeight() <= customArea.getHeight() + 0.5,
            "table should not overflow the bottom edge of the custom tableArea");
    }

    @Test
    void templateSlideIsReusedAsTheBaseForTheOutputInsteadOfABlankSlide() throws Exception {
        var templateBytes = buildTemplateWithLogo();
        var tableArea = new Rectangle2D.Double(24, 60, 400, 300);

        byte[] pptxBytes;
        try (var excelInput = Files.newInputStream(xlsxPath);
             var templateInput = new ByteArrayInputStream(templateBytes)) {
            var params = ConvertParams.builder()
                .excelInput(excelInput)
                .sheetIndex(0)
                .templateInput(templateInput)
                .tableArea(tableArea)
                .build();
            pptxBytes = new SpreadsheetToPptx().convert(params);
        }
        Files.write(pptxPath, pptxBytes);

        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            assertEquals(1, ppt.getSlides().size(),
                "the template's single slide should be reused in place, not appended as a new one");

            var slide = ppt.getSlides().get(0);
            var hasLogo = slide.getShapes().stream()
                .filter(XSLFTextBox.class::isInstance)
                .map(shape -> ((XSLFTextBox) shape).getText())
                .anyMatch("LOGO"::equals);
            assertTrue(hasLogo, "the template's logo textbox should survive into the output");

            var table = (XSLFTable) slide.getShapes().stream()
                .filter(XSLFTable.class::isInstance)
                .findFirst()
                .orElseThrow();
            assertEquals(tableArea.getY(), table.getAnchor().getY(), 0.5,
                "the table should be placed at the custom tableArea's position on the template slide");
        }
    }

    // A minimal single-slide template carrying a "logo" textbox, standing in for a real branded
    // template a caller might supply.
    private static byte[] buildTemplateWithLogo() throws Exception {
        try (var ppt = new XMLSlideShow()) {
            var slide = ppt.createSlide();
            var logo = slide.createTextBox();
            logo.setText("LOGO");
            logo.setAnchor(new Rectangle2D.Double(24, 10, 100, 30));

            var out = new ByteArrayOutputStream();
            ppt.write(out);
            return out.toByteArray();
        }
    }

    @Test
    void everyRenderedFontSizeIsPositiveAndNoLargerThanTheSheetsBaseFont() throws Exception {
        convertFixture();

        var table = readBackTable();

        for (var row : table.getRows()) {
            for (var cell : row.getCells()) {
                for (var para : cell.getTextParagraphs()) {
                    for (var run : para.getTextRuns()) {
                        if (run.getRawText() == null || run.getRawText().isEmpty()) continue;
                        assertTrue(run.getFontSize() > 0, "font size must be positive");
                        assertTrue(run.getFontSize() <= 11.0, "font size must not exceed the sheet's base font");
                    }
                }
            }
        }
    }

    private void convertFixture() throws Exception {
        convertFixture(UnaryOperator.identity());
    }

    // Opens the fixture .xlsx, runs it through the real convert() pipeline, and writes the
    // returned bytes back out to pptxPath -- mirroring what a real caller (e.g. SpreadsheetToPptx's
    // own main()) is now responsible for doing itself.
    private void convertFixture(UnaryOperator<ConvertParams.ConvertParamsBuilder> customizer) throws Exception {
        byte[] pptxBytes;
        try (var excelInput = Files.newInputStream(xlsxPath)) {
            var params = customizer.apply(ConvertParams.builder().excelInput(excelInput).sheetIndex(0)).build();
            pptxBytes = new SpreadsheetToPptx().convert(params);
        }
        Files.write(pptxPath, pptxBytes);
    }

    // Deliberately does not close the XMLSlideShow -- the returned XSLFTable wraps XML objects
    // owned by it, so it's kept open (and closed in @AfterEach) for the rest of the test instead.
    private XSLFTable readBackTable() throws Exception {
        openedResult = new XMLSlideShow(Files.newInputStream(pptxPath));
        var slide = openedResult.getSlides().get(0);
        return (XSLFTable) slide.getShapes().stream()
            .filter(XSLFTable.class::isInstance)
            .findFirst()
            .orElseThrow();
    }
}
