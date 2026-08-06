package xlsxtopptx;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** End-to-end tests for {@link PresentationBuilder}: the {@code %_SHEET(...)} table-replacement
 *  pass and the {@code $NAME} placeholder-substitution pass, run against real in-memory
 *  .pptx/.xlsx bytes the same way {@link SpreadsheetToPptxAcceptanceTest} exercises
 *  {@link SpreadsheetToPptx}. */
class PresentationBuilderTest {

    @Test
    void wholeSheetDirectiveReplacesTheShapeWithAFullUsedRangeTable() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "%_SHEET(Data)", 50, 60, 400, 300));

        try (var ppt = new XMLSlideShow(new ByteArrayInputStream(build(templateBytes, buildWorkbook())))) {
            assertEquals(1, ppt.getSlides().size());
            var slide = ppt.getSlides().get(0);

            var table = tableOn(slide);
            assertEquals(3, table.getNumberOfRows());
            assertEquals(3, table.getNumberOfColumns());
            assertEquals("Alice", table.getCell(1, 0).getText());
            assertEquals("Bob", table.getCell(2, 0).getText());

            var anchor = table.getAnchor();
            assertEquals(60.0, anchor.getY(), 0.5, "table should sit at the top of the placeholder's own bounds");
            assertEquals(400.0, anchor.getWidth(), 0.5, "table should stretch to exactly fill the placeholder's width");
            assertEquals(300.0, anchor.getHeight(), 0.5, "table should stretch to exactly fill the placeholder's height");

            assertFalse(slide.getShapes().stream().anyMatch(XSLFTextBox.class::isInstance),
                "the placeholder shape itself should have been removed");
        }
    }

    @Test
    void rangedDirectiveReplacesTheShapeWithJustThatRange() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "%_SHEET(Data!A1:B2)", 10, 10, 200, 100));

        try (var ppt = new XMLSlideShow(new ByteArrayInputStream(build(templateBytes, buildWorkbook())))) {
            var table = tableOn(ppt.getSlides().get(0));

            assertEquals(2, table.getNumberOfRows());
            assertEquals(2, table.getNumberOfColumns());
            assertEquals("Name", table.getCell(0, 0).getText());
            assertEquals("Alice", table.getCell(1, 0).getText());
        }
    }

    @Test
    void placeholdersOnDifferentSlidesResolveIndependently() throws Exception {
        var templateBytes = buildTemplate(ppt -> {
            addTextBox(ppt.createSlide(), "%_SHEET(Data)", 10, 10, 300, 200);
            addTextBox(ppt.createSlide(), "%_SHEET(Other)", 10, 10, 300, 200);
        });

        try (var ppt = new XMLSlideShow(new ByteArrayInputStream(build(templateBytes, buildWorkbook())))) {
            assertEquals(2, ppt.getSlides().size());

            var firstTable = tableOn(ppt.getSlides().get(0));
            assertEquals(3, firstTable.getNumberOfRows());
            assertEquals("Alice", firstTable.getCell(1, 0).getText());

            var secondTable = tableOn(ppt.getSlides().get(1));
            assertEquals(2, secondTable.getNumberOfRows());
            assertEquals("X", secondTable.getCell(0, 0).getText());
        }
    }

    @Test
    void aSlideWithNoDirectiveIsLeftUntouched() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "Just some text", 10, 10, 200, 50));

        try (var ppt = new XMLSlideShow(new ByteArrayInputStream(build(templateBytes, buildWorkbook())))) {
            var slide = ppt.getSlides().get(0);
            assertFalse(slide.getShapes().stream().anyMatch(XSLFTable.class::isInstance),
                "no table should be created on a slide with no directive");

            var textBox = (XSLFTextBox) slide.getShapes().stream()
                .filter(XSLFTextBox.class::isInstance).findFirst().orElseThrow();
            assertEquals("Just some text", textBox.getText());
        }
    }

    @Test
    void variableTokenIsSubstitutedWithItsRegisteredValue() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "Report for $CLIENT_NAME", 10, 10, 200, 50));

        byte[] result;
        try (var templateIn = new ByteArrayInputStream(templateBytes);
             var excelIn = new ByteArrayInputStream(buildWorkbook())) {
            result = PresentationBuilder.of(templateIn, excelIn)
                .placeholder("CLIENT_NAME", "Acme Corp")
                .build();
        }

        try (var ppt = new XMLSlideShow(new ByteArrayInputStream(result))) {
            var textBox = (XSLFTextBox) ppt.getSlides().get(0).getShapes().stream()
                .filter(XSLFTextBox.class::isInstance).findFirst().orElseThrow();
            assertEquals("Report for Acme Corp", textBox.getText());
        }
    }

    @Test
    void unknownTokenWithNoRegisteredPlaceholderThrows() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "$MYSTERY", 10, 10, 200, 50));

        var ex = assertThrows(IllegalArgumentException.class, () -> build(templateBytes, buildWorkbook()));
        assertTrue(ex.getMessage().contains("MYSTERY"));
    }

    @Test
    void registeredPlaceholderNeverReferencedThrows() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "no tokens here", 10, 10, 200, 50));

        var ex = assertThrows(IllegalArgumentException.class, () -> {
            try (var templateIn = new ByteArrayInputStream(templateBytes);
                 var excelIn = new ByteArrayInputStream(buildWorkbook())) {
                PresentationBuilder.of(templateIn, excelIn).placeholder("UNUSED", "x").build();
            }
        });
        assertTrue(ex.getMessage().contains("UNUSED"));
    }

    @Test
    void malformedSheetDirectiveThrows() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "%_SHEET(Data", 10, 10, 200, 50));

        var ex = assertThrows(IllegalArgumentException.class, () -> build(templateBytes, buildWorkbook()));
        assertTrue(ex.getMessage().contains("Malformed"));
    }

    @Test
    void unknownSheetNameThrows() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "%_SHEET(NoSuchSheet)", 10, 10, 200, 50));

        var ex = assertThrows(IllegalArgumentException.class, () -> build(templateBytes, buildWorkbook()));
        assertTrue(ex.getMessage().contains("NoSuchSheet"));
    }

    @Test
    void invalidRangeThrows() throws Exception {
        var templateBytes = buildTemplate(ppt -> addTextBox(ppt.createSlide(), "%_SHEET(Data!@@@)", 10, 10, 200, 50));

        var ex = assertThrows(IllegalArgumentException.class, () -> build(templateBytes, buildWorkbook()));
        assertTrue(ex.getMessage().contains("@@@"));
    }

    @Test
    void tokenSplitAcrossTwoRunsThrows() throws Exception {
        var templateBytes = buildTemplate(ppt -> {
            var slide = ppt.createSlide();
            var box = slide.createTextBox();
            box.setAnchor(new Rectangle2D.Double(10, 10, 200, 50));
            var para = box.addNewTextParagraph();
            para.addNewTextRun().setText("Value: $F");
            para.addNewTextRun().setText("OO end");
        });

        var ex = assertThrows(IllegalArgumentException.class, () -> {
            try (var templateIn = new ByteArrayInputStream(templateBytes);
                 var excelIn = new ByteArrayInputStream(buildWorkbook())) {
                PresentationBuilder.of(templateIn, excelIn).placeholder("FOO", "filled").build();
            }
        });
        assertTrue(ex.getMessage().contains("split"));
    }

    // ---------- fixtures ----------

    private static byte[] build(byte[] templateBytes, byte[] workbookBytes) throws IOException {
        try (var templateIn = new ByteArrayInputStream(templateBytes);
             var excelIn = new ByteArrayInputStream(workbookBytes)) {
            return PresentationBuilder.of(templateIn, excelIn).build();
        }
    }

    private static byte[] buildTemplate(Consumer<XMLSlideShow> customizer) throws IOException {
        try (var ppt = new XMLSlideShow()) {
            customizer.accept(ppt);
            var out = new ByteArrayOutputStream();
            ppt.write(out);
            return out.toByteArray();
        }
    }

    private static void addTextBox(XSLFSlide slide, String text, double x, double y, double w, double h) {
        var box = slide.createTextBox();
        box.setText(text);
        box.setAnchor(new Rectangle2D.Double(x, y, w, h));
    }

    // "Data": Name/Q1/Q2 header plus two rows (3x3 used range). "Other": a small unrelated 2x2
    // sheet, so the multi-slide test can tell the two placeholders apart.
    private static byte[] buildWorkbook() throws IOException {
        try (var wb = new XSSFWorkbook()) {
            var data = wb.createSheet("Data");
            writeRow(data, 0, "Name", "Q1", "Q2");
            writeRow(data, 1, "Alice", "10", "20");
            writeRow(data, 2, "Bob", "30", "40");

            var other = wb.createSheet("Other");
            writeRow(other, 0, "X", "Y");
            writeRow(other, 1, "1", "2");

            var out = new ByteArrayOutputStream();
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static void writeRow(Sheet sheet, int rowIdx, String... values) {
        var row = sheet.createRow(rowIdx);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    private static XSLFTable tableOn(XSLFSlide slide) {
        return (XSLFTable) slide.getShapes().stream()
            .filter(XSLFTable.class::isInstance)
            .findFirst()
            .orElseThrow();
    }
}
