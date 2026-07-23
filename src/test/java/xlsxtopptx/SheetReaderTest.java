package xlsxtopptx;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.Color;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class SheetReaderTest {

    @Test
    void readsPlainTextAndAppliesDataFormatting() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("Hello");

            var numericCell = row.createCell(1);
            numericCell.setCellValue(1234.5);
            var style = workbook.createCellStyle();
            style.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
            numericCell.setCellStyle(style);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals("Hello", snapshot.rows().get(0).cellAt(0).text());
            assertEquals("1,234.50", snapshot.rows().get(0).cellAt(1).text());
        }
    }

    @Test
    void readsFontAttributes() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            var cell = row.createCell(0);
            cell.setCellValue("Styled");

            var font = workbook.createFont();
            font.setBold(true);
            font.setItalic(true);
            font.setFontHeightInPoints((short) 18);
            font.setFontName("Arial");
            var style = workbook.createCellStyle();
            style.setFont(font);
            cell.setCellStyle(style);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertTrue(cs.bold());
            assertTrue(cs.italic());
            assertEquals(18.0, cs.fontSize());
            assertEquals("Arial", cs.fontName());
        }
    }

    @Test
    void readsSolidFillColor() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            var cell = row.createCell(0);
            cell.setCellValue("Filled");

            var style = (XSSFCellStyle) workbook.createCellStyle();
            style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0x1F, (byte) 0x1F, (byte) 0x6E}));
            style.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertEquals(new Color(0x1F, 0x1F, 0x6E), cs.fill());
        }
    }

    @Test
    void readsBorderEvenWhenApplyBorderFlagIsMissing() throws Exception {
        // Regression test: XSSFCellStyle.getBorderBottom() returns NONE unless the underlying
        // <xf> has an explicit applyBorder="1" attribute. Real-world writers like openpyxl define
        // a genuine border but never emit that flag (Excel doesn't require it), so SheetReader
        // must force it on before reading -- otherwise borders from many real XLSX files vanish.
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            var cell = row.createCell(0);
            cell.setCellValue("Bordered");

            var style = (XSSFCellStyle) workbook.createCellStyle();
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(new XSSFColor(new byte[]{(byte) 0xAD, (byte) 0xD8, (byte) 0xE6}));
            cell.setCellStyle(style);

            // Simulate a writer that never sets applyBorder, regardless of what POI's own setters did.
            style.getCoreXf().unsetApplyBorder();
            assertFalse(style.getCoreXf().getApplyBorder(), "test setup: applyBorder must be unset here");

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertTrue(cs.borders().bottom().isVisible());
            assertEquals(BorderStyle.THIN, cs.borders().bottom().style());
            assertEquals(new Color(0xAD, 0xD8, 0xE6), cs.borders().bottom().color());
        }
    }

    @Test
    void cellsWithoutABorderReportNoneOnEveryEdge() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("Plain");

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertFalse(cs.borders().top().isVisible());
            assertFalse(cs.borders().bottom().isVisible());
            assertFalse(cs.borders().left().isVisible());
            assertFalse(cs.borders().right().isVisible());
        }
    }

    @Test
    void mapsHorizontalAlignment() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);

            var centered = row.createCell(0);
            centered.setCellValue("c");
            var centerStyle = workbook.createCellStyle();
            centerStyle.setAlignment(HorizontalAlignment.CENTER);
            centered.setCellStyle(centerStyle);

            var right = row.createCell(1);
            right.setCellValue("r");
            var rightStyle = workbook.createCellStyle();
            rightStyle.setAlignment(HorizontalAlignment.RIGHT);
            right.setCellStyle(rightStyle);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals(TextAlign.CENTER, snapshot.rows().get(0).cellAt(0).align());
            assertEquals(TextAlign.RIGHT, snapshot.rows().get(0).cellAt(1).align());
        }
    }

    @Test
    void readsExplicitMergedRegions() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            sheet.createRow(0).createCell(3).setCellValue("Banner");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 3, 4));

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals(1, snapshot.merges().size());
            assertEquals(new MergeRegion(0, 0, 3, 4), snapshot.merges().get(0));
        }
    }

    @Test
    void detectsCenterAcrossSelectionAsAMergeLookAlike() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            var centerAcrossStyle = workbook.createCellStyle();
            centerAcrossStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);

            var first = row.createCell(1);
            first.setCellValue("Spans");
            first.setCellStyle(centerAcrossStyle);
            var second = row.createCell(2);
            second.setCellStyle(centerAcrossStyle);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals(1, snapshot.merges().size());
            assertEquals(new MergeRegion(0, 0, 1, 2), snapshot.merges().get(0));
        }
    }

    @Test
    void aSingleCenterAcrossSelectionCellDoesNotCountAsAMerge() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            var style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            var cell = row.createCell(0);
            cell.setCellValue("Alone");
            cell.setCellStyle(style);

            var snapshot = SheetReader.read(sheet, 0);

            assertTrue(snapshot.merges().isEmpty());
        }
    }

    @Test
    void blankCellsInARowAreReadAsEmptySnapshots() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(2).setCellValue("only this one");

            var snapshot = SheetReader.read(sheet, 0);

            assertTrue(snapshot.rows().get(0).cellAt(0).isEmpty());
            assertTrue(snapshot.rows().get(0).cellAt(1).isEmpty());
            assertEquals("only this one", snapshot.rows().get(0).cellAt(2).text());
        }
    }
}
