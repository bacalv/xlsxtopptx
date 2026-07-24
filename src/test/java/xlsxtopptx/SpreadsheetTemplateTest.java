package xlsxtopptx;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A minimal stand-in for the real test_sheet_template.xlsx this class was built against: three
 * marker rows ("%_BOLD", "%_BLANK", "%_PLAIN") each with two "%_0"/"%_1" placeholders and a
 * per-row SUM formula, followed by a grand-total row that sums each column via a
 * COL_B_VALUES/COL_C_VALUES named range instead of a hardcoded cell range.
 */
class SpreadsheetTemplateTest {
    private static final String SHEET = "Sheet1";

    @Test
    void substitutesLabelAndPlaceholdersPerRow() throws Exception {
        var rendered = render(buildTemplate(false), List.of(
            SpreadsheetRow.type("BOLD").data("Migrate billing", 10, 20).build(),
            SpreadsheetRow.type("PLAIN").data("Retire legacy auth", 5, 7).build()));

        var sheet = rendered.getSheetAt(0);
        assertEquals("Migrate billing", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals(10.0, sheet.getRow(0).getCell(1).getNumericCellValue());
        assertEquals(20.0, sheet.getRow(0).getCell(2).getNumericCellValue());
        assertEquals("Retire legacy auth", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals(5.0, sheet.getRow(1).getCell(1).getNumericCellValue());
        assertEquals(7.0, sheet.getRow(1).getCell(2).getNumericCellValue());
    }

    @Test
    void blankTypeNeedsNoDataCall() throws Exception {
        var rendered = render(buildTemplate(false), List.of(
            SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build(),
            SpreadsheetRow.type("BLANK").build(),
            SpreadsheetRow.type("PLAIN").data("Row B", 3, 4).build()));

        var sheet = rendered.getSheetAt(0);
        assertEquals("", sheet.getRow(1).getCell(0).getStringCellValue());
    }

    @Test
    void perRowFormulaIsShiftedToTheOutputRowNumber() throws Exception {
        var rendered = render(buildTemplate(false), List.of(
            SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build(),
            SpreadsheetRow.type("BLANK").build(),
            SpreadsheetRow.type("PLAIN").data("Row B", 10, 20).build()));

        var sheet = rendered.getSheetAt(0);
        // "Row B" is rendered as the 3rd output row -- its per-row formula (originally SUM(B3:C3)
        // on the template's own 3rd marker row) must still reference its own row, wherever that
        // output row actually lands.
        assertEquals("SUM(B3:C3)", sheet.getRow(2).getCell(3).getCellFormula());
        var evaluator = rendered.getCreationHelper().createFormulaEvaluator();
        assertEquals(30.0, evaluator.evaluate(sheet.getRow(2).getCell(3)).getNumberValue());
    }

    @Test
    void namedRangeGrowsToCoverMoreRowsThanTheOriginalTemplateBlock() throws Exception {
        var rendered = render(buildTemplate(false), List.of(
            SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build(),
            SpreadsheetRow.type("PLAIN").data("Row B", 3, 4).build(),
            SpreadsheetRow.type("PLAIN").data("Row C", 5, 6).build(),
            SpreadsheetRow.type("PLAIN").data("Row D", 7, 8).build()));

        assertEquals(SHEET + "!$B$1:$B$4", rendered.getName("COL_B_VALUES").getRefersToFormula());
        assertEquals(SHEET + "!$C$1:$C$4", rendered.getName("COL_C_VALUES").getRefersToFormula());

        var sheet = rendered.getSheetAt(0);
        var totalRow = sheet.getRow(4); // moved down from its original row index 3
        assertEquals("Total", totalRow.getCell(0).getStringCellValue());

        var evaluator = rendered.getCreationHelper().createFormulaEvaluator();
        assertEquals(16.0, evaluator.evaluate(totalRow.getCell(1)).getNumberValue()); // 1+3+5+7
        assertEquals(20.0, evaluator.evaluate(totalRow.getCell(2)).getNumberValue()); // 2+4+6+8
    }

    @Test
    void namedRangeShrinksWhenFewerRowsAreRendered() throws Exception {
        var rendered = render(buildTemplate(false), List.of(
            SpreadsheetRow.type("BOLD").data("Only row", 9, 11).build()));

        assertEquals(SHEET + "!$B$1:$B$1", rendered.getName("COL_B_VALUES").getRefersToFormula());

        var sheet = rendered.getSheetAt(0);
        var totalRow = sheet.getRow(1); // moved up from its original row index 3
        assertEquals("Total", totalRow.getCell(0).getStringCellValue());
        var evaluator = rendered.getCreationHelper().createFormulaEvaluator();
        assertEquals(9.0, evaluator.evaluate(totalRow.getCell(1)).getNumberValue());
    }

    @Test
    void survivesAGenuineExcelSharedFormulaGroupInTheTotalRow() throws Exception {
        // Regression test: a template built by hand in Excel/Sheets -- as the real
        // test_sheet_template.xlsx was -- is likely to store a dragged-across formula row as a
        // "shared formula" group. Sheet.shiftRows re-derives shared-formula siblings from their
        // group's master cell by column offset, which collapses every sibling onto the master's
        // exact text for a named-range formula (nothing column-relative to shift), silently
        // corrupting every column but the first. render() must de-share those cells first.
        var rendered = render(buildTemplate(true), List.of(
            SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build(),
            SpreadsheetRow.type("PLAIN").data("Row B", 3, 4).build(),
            SpreadsheetRow.type("PLAIN").data("Row C", 5, 6).build()));

        var sheet = rendered.getSheetAt(0);
        var totalRow = sheet.getRow(3);
        assertEquals("SUM(COL_B_VALUES)", totalRow.getCell(1).getCellFormula());
        assertEquals("SUM(COL_C_VALUES)", totalRow.getCell(2).getCellFormula());

        var evaluator = rendered.getCreationHelper().createFormulaEvaluator();
        assertEquals(9.0, evaluator.evaluate(totalRow.getCell(1)).getNumberValue());
        assertEquals(12.0, evaluator.evaluate(totalRow.getCell(2)).getNumberValue());
    }

    @Test
    void boldStyleIsPreservedPerType() throws Exception {
        var rendered = render(buildTemplate(false), List.of(
            SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build(),
            SpreadsheetRow.type("PLAIN").data("Row B", 3, 4).build()));

        var sheet = rendered.getSheetAt(0);
        assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getFont().getBold());
        assertFalse(sheet.getRow(1).getCell(0).getCellStyle().getFont().getBold());
    }

    @Test
    void unknownTypeThrows() throws IOException {
        var template = buildTemplate(false);
        var ex = assertThrows(IllegalArgumentException.class, () -> SpreadsheetTemplate.of(new ByteArrayInputStream(template))
            .render(List.of(SpreadsheetRow.type("NOPE").build())));
        assertTrue(ex.getMessage().contains("NOPE"));
    }

    @Test
    void mismatchedPlaceholderCountThrows() throws IOException {
        var template = buildTemplate(false);
        assertThrows(IllegalArgumentException.class, () -> SpreadsheetTemplate.of(new ByteArrayInputStream(template))
            .render(List.of(SpreadsheetRow.type("BOLD").data("Row A", 1).build())));
    }

    @Test
    void emptyRowListThrows() throws IOException {
        var template = buildTemplate(false);
        assertThrows(IllegalArgumentException.class, () -> SpreadsheetTemplate.of(new ByteArrayInputStream(template))
            .render(List.of()));
    }

    @Test
    void duplicateTypeMarkerThrows() throws IOException {
        var template = buildTemplateWithDuplicateMarker();
        var ex = assertThrows(IllegalStateException.class, () -> SpreadsheetTemplate.of(new ByteArrayInputStream(template))
            .render(List.of(SpreadsheetRow.type("PLAIN").data("x", 1, 2).build())));
        assertTrue(ex.getMessage().contains("PLAIN"));
    }

    @Test
    void placeholderTokenIsSubstitutedWhereverItAppearsInTheSheet() throws Exception {
        var template = buildTemplateWithTitleCell("As of $REPORT_DATE");
        var rendered = new XSSFWorkbook(new ByteArrayInputStream(SpreadsheetTemplate.of(new ByteArrayInputStream(template))
            .placeholder("REPORT_DATE", "2026-06-11")
            .render(List.of(SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build()))));

        assertEquals("As of 2026-06-11", rendered.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
    }

    @Test
    void unknownPlaceholderTokenThrows() throws IOException {
        var template = buildTemplateWithTitleCell("As of $REPORT_DATE");
        var ex = assertThrows(IllegalArgumentException.class,
            () -> SpreadsheetTemplate.of(new ByteArrayInputStream(template))
                .render(List.of(SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build())));
        assertTrue(ex.getMessage().contains("REPORT_DATE"));
    }

    @Test
    void unusedRegisteredPlaceholderThrows() throws IOException {
        var template = buildTemplate(false);
        var ex = assertThrows(IllegalArgumentException.class,
            () -> SpreadsheetTemplate.of(new ByteArrayInputStream(template))
                .placeholder("REPORT_DATE", "2026-06-11")
                .render(List.of(SpreadsheetRow.type("BOLD").data("Row A", 1, 2).build())));
        assertTrue(ex.getMessage().contains("REPORT_DATE"));
    }

    private static XSSFWorkbook render(byte[] template, List<SpreadsheetRow> rows) throws IOException {
        var rendered = SpreadsheetTemplate.of(new ByteArrayInputStream(template)).render(rows);
        return new XSSFWorkbook(new ByteArrayInputStream(rendered));
    }

    private static byte[] buildTemplate(boolean sharedFormulaTotalRow) throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(SHEET);

            var boldFont = workbook.createFont();
            boldFont.setBold(true);
            var boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);

            markerRow(sheet, 0, "%_BOLD", boldStyle, "SUM(B1:C1)");
            var blankRow = sheet.createRow(1);
            blankRow.createCell(0).setCellValue("%_BLANK");
            markerRow(sheet, 2, "%_PLAIN", null, "SUM(B3:C3)");

            var nameB = workbook.createName();
            nameB.setNameName("COL_B_VALUES");
            nameB.setRefersToFormula(SHEET + "!$B$1:$B$3");
            var nameC = workbook.createName();
            nameC.setNameName("COL_C_VALUES");
            nameC.setRefersToFormula(SHEET + "!$C$1:$C$3");

            var totalRow = sheet.createRow(3);
            totalRow.createCell(0).setCellValue("Total");
            var totalB = totalRow.createCell(1);
            totalB.setCellFormula("SUM(COL_B_VALUES)");
            var totalC = totalRow.createCell(2);
            totalC.setCellFormula("SUM(COL_C_VALUES)");

            if (sharedFormulaTotalRow) {
                shareFormulaGroup(totalB, totalC, "B4:C4");
            }

            var out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] buildTemplateWithTitleCell(String titleText) throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(SHEET);
            markerRow(sheet, 0, "%_BOLD", null, "SUM(B1:C1)");
            sheet.createRow(2).createCell(0).setCellValue(titleText);

            var out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] buildTemplateWithDuplicateMarker() throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(SHEET);
            markerRow(sheet, 0, "%_PLAIN", null, "SUM(B1:C1)");
            markerRow(sheet, 1, "%_PLAIN", null, "SUM(B2:C2)");

            var out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static void markerRow(org.apache.poi.xssf.usermodel.XSSFSheet sheet, int rowIndex, String marker,
                                   org.apache.poi.xssf.usermodel.XSSFCellStyle style, String formula) {
        var row = sheet.createRow(rowIndex);
        var label = row.createCell(0);
        label.setCellValue(marker);
        if (style != null) label.setCellStyle(style);

        var placeholder0 = row.createCell(1);
        placeholder0.setCellValue("%_0");
        if (style != null) placeholder0.setCellStyle(style);

        var placeholder1 = row.createCell(2);
        placeholder1.setCellValue("%_1");
        if (style != null) placeholder1.setCellStyle(style);

        row.createCell(3).setCellFormula(formula);
    }

    /** Mimics what Excel/Sheets writes when a user drags a formula across a row: a "master" cell
     *  carrying the shared group's shape plus each cell's own resolved formula text. */
    private static void shareFormulaGroup(org.apache.poi.ss.usermodel.Cell master, org.apache.poi.ss.usermodel.Cell sibling, String ref) {
        var masterFormula = ((org.apache.poi.xssf.usermodel.XSSFCell) master).getCTCell().getF();
        masterFormula.setT(STCellFormulaType.SHARED);
        masterFormula.setRef(ref);
        masterFormula.setSi(1);

        var siblingFormula = ((org.apache.poi.xssf.usermodel.XSSFCell) sibling).getCTCell().getF();
        siblingFormula.setT(STCellFormulaType.SHARED);
        siblingFormula.setSi(1);
    }
}
