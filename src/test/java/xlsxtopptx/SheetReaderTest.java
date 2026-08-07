package xlsxtopptx;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.CTColor;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeDocument;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;

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
    void resolvesThemedFillColorIncludingItsTint() throws Exception {
        // Regression test: colors picked from Excel's "Theme Colors" palette (as opposed to
        // "Standard Colors" or a custom RGB) are stored as a theme index plus a tint, not as an
        // RGB value -- e.g. the light banded-row blue in a styled table is typically
        // "Blue, Accent 1, Lighter 60%". XSSFColor.getRGB() only returns an explicitly-stored RGB,
        // so without resolving against the workbook's theme and reapplying the tint, this fill
        // silently disappears from the rendered slide.
        try (var workbook = workbookWithTheme()) {
            var sheet = workbook.createSheet();
            var cell = sheet.createRow(0).createCell(0);
            cell.setCellValue("Banded");

            var style = (XSSFCellStyle) workbook.createCellStyle();
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            var themedColor = new XSSFColor();
            themedColor.setTheme(4); // accent1, set to 0x4F81BD below
            themedColor.setTint(0.6); // "Lighter 60%"
            style.setFillForegroundColor(themedColor);
            cell.setCellStyle(style);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertEquals(new Color(0xB8, 0xCC, 0xE4), cs.fill());
        }
    }

    /** Builds a workbook with a minimal, real theme part -- {@code new XSSFWorkbook()} alone has
     *  none, but every genuine Excel- or openpyxl-authored file does, since a theme is how "Theme
     *  Color" fills/fonts (as opposed to standard/custom RGB ones) are resolved. */
    private static XSSFWorkbook workbookWithTheme() {
        var themeDoc = ThemeDocument.Factory.newInstance();
        var scheme = themeDoc.addNewTheme().addNewThemeElements().addNewClrScheme();
        setSrgb(scheme.addNewDk1(), 0, 0, 0);
        setSrgb(scheme.addNewLt1(), 255, 255, 255);
        setSrgb(scheme.addNewDk2(), 0, 0, 0);
        setSrgb(scheme.addNewLt2(), 255, 255, 255);
        setSrgb(scheme.addNewAccent1(), 0x4F, 0x81, 0xBD);
        setSrgb(scheme.addNewAccent2(), 0, 0, 0);
        setSrgb(scheme.addNewAccent3(), 0, 0, 0);
        setSrgb(scheme.addNewAccent4(), 0, 0, 0);
        setSrgb(scheme.addNewAccent5(), 0, 0, 0);
        setSrgb(scheme.addNewAccent6(), 0, 0, 0);
        setSrgb(scheme.addNewHlink(), 0, 0, 0);
        setSrgb(scheme.addNewFolHlink(), 0, 0, 0);

        var workbook = new XSSFWorkbook();
        workbook.getStylesSource().setTheme(new ThemesTable(themeDoc));
        return workbook;
    }

    private static void setSrgb(CTColor color, int r, int g, int b) {
        color.addNewSrgbClr().setVal(new byte[]{(byte) r, (byte) g, (byte) b});
    }

    @Test
    void appliesConditionalFormattingFillWhenTheRuleConditionIsTrue() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue(-5);

            addFillRule(sheet, "A1:A1", ComparisonOperator.LT, "0", 0xFF, 0x00, 0x00);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertEquals(new Color(0xFF, 0x00, 0x00), cs.fill());
        }
    }

    @Test
    void conditionalFormattingLeavesFillUnchangedWhenTheRuleConditionIsFalse() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue(5);

            addFillRule(sheet, "A1:A1", ComparisonOperator.LT, "0", 0xFF, 0x00, 0x00);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertEquals(null, cs.fill());
        }
    }

    @Test
    void conditionalFormattingFillOverridesTheCellsOwnStaticFill() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            var cell = row.createCell(0);
            cell.setCellValue(-5);

            var style = (XSSFCellStyle) workbook.createCellStyle();
            style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0x00, (byte) 0x00, (byte) 0xFF}));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);

            addFillRule(sheet, "A1:A1", ComparisonOperator.LT, "0", 0xFF, 0x00, 0x00);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertEquals(new Color(0xFF, 0x00, 0x00), cs.fill());
        }
    }

    @Test
    void higherPriorityConditionalFormattingRuleWinsWhenMultipleRulesMatch() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue(-5);

            // Both rules match (-5 < 0 and -5 <= 10); the first-added one has the higher priority
            // (lower priority number) in XSSF, so its fill should win.
            addFillRule(sheet, "A1:A1", ComparisonOperator.LT, "0", 0xFF, 0x00, 0x00);
            addFillRule(sheet, "A1:A1", ComparisonOperator.LE, "10", 0x00, 0xFF, 0x00);

            var cs = SheetReader.read(sheet, 0).rows().get(0).cellAt(0);

            assertEquals(new Color(0xFF, 0x00, 0x00), cs.fill());
        }
    }

    @Test
    void redIfNegativeGreenOtherwiseWorksAcrossMultipleRows() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            sheet.createRow(0).createCell(0).setCellValue(-10);
            sheet.createRow(1).createCell(0).setCellValue(0);
            sheet.createRow(2).createCell(0).setCellValue(10);

            addFillRule(sheet, "A1:A3", ComparisonOperator.LT, "0", 0xFF, 0x00, 0x00);
            addFillRule(sheet, "A1:A3", ComparisonOperator.GE, "0", 0x00, 0xFF, 0x00);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals(new Color(0xFF, 0x00, 0x00), snapshot.rows().get(0).cellAt(0).fill());
            assertEquals(new Color(0x00, 0xFF, 0x00), snapshot.rows().get(1).cellAt(0).fill());
            assertEquals(new Color(0x00, 0xFF, 0x00), snapshot.rows().get(2).cellAt(0).fill());
        }
    }

    private static void addFillRule(
        org.apache.poi.ss.usermodel.Sheet sheet, String range, byte comparisonOperator, String formula1,
        int r, int g, int b
    ) {
        var sheetCf = sheet.getSheetConditionalFormatting();
        var rule = sheetCf.createConditionalFormattingRule(comparisonOperator, formula1);
        var fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b}));
        fill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        sheetCf.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf(range)}, rule);
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
    void generalAlignmentRightAlignsNumbersAndLeftAlignsText() throws Exception {
        // Regression test: a cell with no explicit alignment override reports
        // HorizontalAlignment.GENERAL, which previously always mapped to TextAlign.LEFT. But
        // Excel's actual "General" alignment right-aligns numbers/dates and only left-aligns text
        // -- so unstyled numeric columns (a very common case) were rendering left-aligned instead
        // of matching the source spreadsheet.
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("text");
            row.createCell(1).setCellValue(1234.5);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals(TextAlign.LEFT, snapshot.rows().get(0).cellAt(0).align());
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

    // Reproduces what Excel's "Break Links" does to a shared formula group when its master cell
    // referenced another workbook: the master is converted to a plain value (its <f> element
    // disappears entirely), but the other cells in the group are untouched and still carry
    // <f t="shared" si="N"/> pointing at a master that's no longer declared anywhere -- which is
    // exactly what makes POI throw "Master cell of a shared formula with sid=N was not found".
    @Test
    void aSharedFormulaGroupWithAMissingMasterFallsBackToEachCellsCachedValue() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");

            var row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue(1);
            var b1 = (XSSFCell) row0.createCell(1);
            b1.setCellFormula("A1*10");

            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(2);
            var b2 = (XSSFCell) row1.createCell(1);
            b2.setCellFormula("A2*10");

            var row2 = sheet.createRow(2);
            row2.createCell(0).setCellValue(3);
            var b3 = (XSSFCell) row2.createCell(1);
            b3.setCellFormula("A3*10");

            // A genuine shared-formula group: B1 is the master (carries ref + formula text), B2
            // and B3 are members that only reference it by si.
            markAsSharedFormulaMaster(b1, "B1:B3", 0);
            markAsOrphanedSharedFormulaMember(b2, 0, "20");
            markAsOrphanedSharedFormulaMember(b3, 0, "30");

            // Simulate "Break Links": the master's formula is stripped to a static value, leaving
            // B2/B3 pointing at a shared-formula group with no master.
            b1.getCTCell().unsetF();
            b1.setCellValue(10.0);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals("10", snapshot.rows().get(0).cellAt(1).text());
            assertEquals("20", snapshot.rows().get(1).cellAt(1).text(), "should fall back to B2's cached value");
            assertEquals("30", snapshot.rows().get(2).cellAt(1).text(), "should fall back to B3's cached value");
        }
    }

    @Test
    void aSharedFormulaGroupMemberWithNoCachedValueFallsBackToBlank() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");

            var row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue(1);
            var b1 = (XSSFCell) row0.createCell(1);
            b1.setCellFormula("A1*10");

            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(2);
            var b2 = (XSSFCell) row1.createCell(1);
            b2.setCellFormula("A2*10");

            markAsSharedFormulaMaster(b1, "B1:B2", 0);
            markAsOrphanedSharedFormulaMember(b2, 0, null); // no cached <v> at all

            b1.getCTCell().unsetF();
            b1.setCellValue(10.0);

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals("", snapshot.rows().get(1).cellAt(1).text());
        }
    }

    // Reproduces a real-world report: an ordinary, perfectly intact formula cell (L7) that merely
    // *references* another cell (D7) whose shared formula group is broken. Evaluating L7 fails
    // deep inside WorkbookEvaluator while resolving that reference, and POI wraps it in an outer
    // "Failed to evaluate cell: ..." IllegalStateException -- the real "Master cell of a shared
    // formula..." failure only shows up via getCause(). L7 itself is untouched by the corruption,
    // so it should recover using its own last-cached value.
    @Test
    void aFormulaThatMerelyReferencesAnotherCellsBrokenSharedFormulaFallsBackToItsOwnCachedValue() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("PS");

            var row6 = sheet.createRow(5); // Excel row 6
            row6.createCell(0).setCellValue(1); // A6
            var d6 = (XSSFCell) row6.createCell(3); // D6: shared-formula master
            d6.setCellFormula("A6*10");

            var row7 = sheet.createRow(6); // Excel row 7
            row7.createCell(0).setCellValue(2); // A7
            var d7 = (XSSFCell) row7.createCell(3); // D7: shared-formula member, si=0
            d7.setCellFormula("A7*10");
            row7.createCell(7).setCellValue(5); // H7
            var l7 = (XSSFCell) row7.createCell(11); // L7: ordinary formula referencing H7 and D7

            markAsSharedFormulaMaster(d6, "D6:D7", 0);
            markAsOrphanedSharedFormulaMember(d7, 0, "20");
            l7.setCellFormula("H7-D7");

            // Break the master, same as Excel's "Break Links" would.
            d6.getCTCell().unsetF();
            d6.setCellValue(10.0);

            // L7's own last-calculated value, as Excel would have cached it before the file broke.
            l7.getCTCell().setV("-45");

            var snapshot = SheetReader.read(sheet, 0);

            assertEquals("-45", snapshot.rows().get(6).cellAt(11).text(),
                "L7 should fall back to its own cached value when evaluating it transitively hits D7's broken group");
        }
    }

    private static void markAsSharedFormulaMaster(XSSFCell cell, String ref, int si) {
        var f = cell.getCTCell().getF();
        f.setT(STCellFormulaType.SHARED);
        f.setRef(ref);
        f.setSi(si);
    }

    private static void markAsOrphanedSharedFormulaMember(XSSFCell cell, int si, String cachedValue) {
        var f = cell.getCTCell().getF();
        f.setT(STCellFormulaType.SHARED);
        f.setSi(si);
        f.setStringValue(""); // no formula text of its own -- relies entirely on the (soon-missing) master
        if (cachedValue != null) {
            cell.getCTCell().setV(cachedValue);
        }
    }
}
