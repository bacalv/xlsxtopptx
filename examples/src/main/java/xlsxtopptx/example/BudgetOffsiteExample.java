package xlsxtopptx.example;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import xlsxtopptx.ConvertParams;
import xlsxtopptx.PptxMerger;
import xlsxtopptx.SpreadsheetRow;
import xlsxtopptx.SpreadsheetTemplate;
import xlsxtopptx.SpreadsheetToPptx;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Generates every file the README's "spreadsheet template to branded slideshow" tutorial walks
 * through, writing them to {@code examples/} for the reader to open, run, or screenshot directly
 * instead of taking the README's word for it. Run via {@code ./gradlew runExample}.
 *
 * <p>The five outputs, in the order this class builds them:
 * <ul>
 *   <li>{@code budget-template.xlsx} -- the hand-authored row template (marker rows, a per-row
 *       variance formula, named-range totals, a {@code $REPORT_DATE} placeholder).
 *   <li>{@code title-slide-template.pptx} / {@code logo-template.pptx} -- two single-slide
 *       branding templates sharing the same masthead and footer, standing in for a real company
 *       deck's title layout and content layout.
 *   <li>{@code budget-rendered.xlsx} -- the template rendered against this month's line items via
 *       {@link SpreadsheetTemplate}.
 *   <li>{@code budget-presentation.pptx} -- the final two-slide deck: the title slide followed by
 *       the rendered budget on the branded content slide, combined with {@link PptxMerger}.
 * </ul>
 */
public final class BudgetOffsiteExample {
    private static final String SHEET_NAME = "Budget";
    private static final String NUMBER_FORMAT = "$#,##0";
    private static final Color NAVY = new Color(0x14, 0x2B, 0x4B);
    private static final Color GOLD = new Color(0xC9, 0xA2, 0x27);
    private static final Color HIGHLIGHT_FILL = new Color(0xFC, 0xE4, 0xD6);
    private static final Color HEADER_FILL = new Color(0xEE, 0xEE, 0xEE);
    private static final double PAGE_WIDTH_PT = 960;
    private static final double PAGE_HEIGHT_PT = 540;
    private static final double MASTHEAD_HEIGHT_PT = 54;

    public static void main(String[] args) throws IOException {
        var outDir = Path.of("examples");
        Files.createDirectories(outDir);

        var templateBytes = buildTemplateWorkbook();
        Files.write(outDir.resolve("budget-template.xlsx"), templateBytes);

        var titleSlideBytes = buildBrandedSlide(true);
        Files.write(outDir.resolve("title-slide-template.pptx"), titleSlideBytes);

        var logoTemplateBytes = buildBrandedSlide(false);
        Files.write(outDir.resolve("logo-template.pptx"), logoTemplateBytes);

        var renderedBytes = SpreadsheetTemplate.of(new ByteArrayInputStream(templateBytes))
            .placeholder("REPORT_DATE", "27 Jul 2026")
            .render(exampleRows());
        Files.write(outDir.resolve("budget-rendered.xlsx"), renderedBytes);

        byte[] dataSlideBytes;
        try (var excelInput = new ByteArrayInputStream(renderedBytes);
             var templateInput = new ByteArrayInputStream(logoTemplateBytes)) {
            var params = ConvertParams.builder()
                .excelInput(excelInput)
                .templateInput(templateInput)
                .tableArea(new Rectangle2D.Double(40, 84, 880, 420))
                .scaleToFitWidth(true)
                .build();
            dataSlideBytes = new SpreadsheetToPptx().convert(params);
        }

        try (var titleInput = new ByteArrayInputStream(titleSlideBytes);
             var dataInput = new ByteArrayInputStream(dataSlideBytes)) {
            var presentationBytes = PptxMerger.merge(List.of(titleInput, dataInput));
            Files.write(outDir.resolve("budget-presentation.pptx"), presentationBytes);
        }

        System.out.println("Wrote budget-template.xlsx, title-slide-template.pptx, logo-template.pptx, "
            + "budget-rendered.xlsx, and budget-presentation.pptx to " + outDir.toAbsolutePath());
    }

    // This month's line items -- swap in a different list (more rows, fewer rows, a different
    // month) and everything downstream -- the marker block, the named-range totals, the variance
    // formulas, the slide layout -- adjusts to fit, which is the whole point of the tutorial.
    private static List<SpreadsheetRow> exampleRows() {
        return List.of(
            SpreadsheetRow.type("HIGHLIGHT").data("Venue hire (The Old Brewery)", 4000, 4250).build(),
            SpreadsheetRow.type("ITEM").data("Catering & drinks", 2200, 2100).build(),
            SpreadsheetRow.type("ITEM").data("Travel & transport", 1800, 1950).build(),
            SpreadsheetRow.type("ITEM").data("Swag & prizes", 600, 580).build(),
            SpreadsheetRow.type("BLANK").build(),
            SpreadsheetRow.type("ITEM").data("AV & entertainment", 900, 900).build());
    }

    /**
     * Builds the row template from scratch with Apache POI directly -- standing in for a sheet a
     * person would normally author by hand in Excel. Layout:
     * <pre>
     * row 1        A1:D1 merged title, "Team Offsite Budget -- as of $REPORT_DATE"
     * row 2        header: Category | Estimated ($) | Actual ($) | Variance ($)
     * row 3        %_HIGHLIGHT marker row (bold, highlighted fill)
     * row 4        %_ITEM marker row (plain)
     * row 5        %_BLANK marker row (spacer, label only)
     * row 6        Total: SUM(COL_EST_VALUES), SUM(COL_ACT_VALUES), and their difference
     * </pre>
     * Rows 3-5 are the contiguous marker block {@link SpreadsheetTemplate#render} requires; the
     * two named ranges span exactly that block so growing or shrinking it (to fit however many
     * rows {@link #exampleRows} supplies) keeps the Total row's formulas correct.
     */
    private static byte[] buildTemplateWorkbook() throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(SHEET_NAME);

            var titleStyle = titleStyle(workbook);
            var headerStyle = headerStyle(workbook);
            var highlightStyle = numberStyle(workbook, true, false);
            var itemStyle = numberStyle(workbook, false, false);
            var totalStyle = numberStyle(workbook, false, true);
            var totalLabelStyle = boldStyle(workbook);

            var titleRow = sheet.createRow(0);
            var titleCell = titleRow.createCell(0);
            titleCell.setCellValue("Team Offsite Budget -- as of $REPORT_DATE");
            titleCell.setCellStyle(titleStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));

            var headerRow = sheet.createRow(1);
            setCell(headerRow, 0, "Category", headerStyle);
            setCell(headerRow, 1, "Estimated ($)", headerStyle);
            setCell(headerRow, 2, "Actual ($)", headerStyle);
            setCell(headerRow, 3, "Variance ($)", headerStyle);

            markerRow(sheet, 2, "%_HIGHLIGHT", highlightStyle, "C3-B3");
            markerRow(sheet, 3, "%_ITEM", itemStyle, "C4-B4");
            sheet.createRow(4).createCell(0).setCellValue("%_BLANK"); // spacer marker, no placeholders

            var estName = workbook.createName();
            estName.setNameName("COL_EST_VALUES");
            estName.setRefersToFormula(SHEET_NAME + "!$B$3:$B$5");
            var actName = workbook.createName();
            actName.setNameName("COL_ACT_VALUES");
            actName.setRefersToFormula(SHEET_NAME + "!$C$3:$C$5");

            var totalRow = sheet.createRow(5);
            setCell(totalRow, 0, "Total", totalLabelStyle);
            var totalEst = totalRow.createCell(1);
            totalEst.setCellFormula("SUM(COL_EST_VALUES)");
            totalEst.setCellStyle(totalStyle);
            var totalAct = totalRow.createCell(2);
            totalAct.setCellFormula("SUM(COL_ACT_VALUES)");
            totalAct.setCellStyle(totalStyle);
            var totalVar = totalRow.createCell(3);
            totalVar.setCellFormula("C6-B6");
            totalVar.setCellStyle(totalStyle);

            sheet.setColumnWidth(0, 34 * 256);
            for (int c = 1; c <= 3; c++) {
                sheet.setColumnWidth(c, 16 * 256);
            }

            var out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static void markerRow(org.apache.poi.xssf.usermodel.XSSFSheet sheet, int rowIndex, String marker,
                                   XSSFCellStyle style, String varianceFormula) {
        var row = sheet.createRow(rowIndex);
        var label = row.createCell(0);
        label.setCellValue(marker);
        label.setCellStyle(style);

        var placeholder0 = row.createCell(1);
        placeholder0.setCellValue("%_0");
        placeholder0.setCellStyle(style);

        var placeholder1 = row.createCell(2);
        placeholder1.setCellValue("%_1");
        placeholder1.setCellStyle(style);

        var variance = row.createCell(3);
        variance.setCellFormula(varianceFormula);
        variance.setCellStyle(style);
    }

    private static void setCell(org.apache.poi.ss.usermodel.Row row, int col, String text, XSSFCellStyle style) {
        var cell = row.createCell(col);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    private static XSSFCellStyle titleStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook));
        style.getFont().setFontHeightInPoints((short) 14);
        return style;
    }

    private static XSSFCellStyle headerStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook));
        style.setFillForegroundColor(new XSSFColor(new byte[]{
            (byte) HEADER_FILL.getRed(), (byte) HEADER_FILL.getGreen(), (byte) HEADER_FILL.getBlue()}));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        return style;
    }

    private static XSSFCellStyle boldStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook));
        style.setBorderTop(BorderStyle.THIN);
        return style;
    }

    private static XSSFCellStyle numberStyle(XSSFWorkbook workbook, boolean highlighted, boolean total) {
        var style = workbook.createCellStyle();
        if (highlighted || total) {
            style.setFont(boldFont(workbook));
        }
        if (highlighted) {
            style.setFillForegroundColor(new XSSFColor(new byte[]{
                (byte) HIGHLIGHT_FILL.getRed(), (byte) HIGHLIGHT_FILL.getGreen(), (byte) HIGHLIGHT_FILL.getBlue()}));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (total) {
            style.setBorderTop(BorderStyle.THIN);
        }
        style.setDataFormat(workbook.createDataFormat().getFormat(NUMBER_FORMAT));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    private static XSSFFont boldFont(XSSFWorkbook workbook) {
        var font = workbook.createFont();
        font.setBold(true);
        return font;
    }

    /**
     * Builds a single-slide branded template from scratch: a navy masthead with a wordmark, a gold
     * footer rule, and -- only for the title slide -- a centered title and subtitle. Both slides
     * share the same masthead/footer so the title slide and the content slide (used later as
     * {@link ConvertParams#getTemplateInput()}) read as one consistent deck once merged.
     */
    private static byte[] buildBrandedSlide(boolean titleSlide) throws IOException {
        try (var ppt = new XMLSlideShow()) {
            ppt.setPageSize(new java.awt.Dimension((int) PAGE_WIDTH_PT, (int) PAGE_HEIGHT_PT));
            var slide = ppt.createSlide();

            var masthead = (XSLFAutoShape) slide.createAutoShape();
            masthead.setShapeType(ShapeType.RECT);
            masthead.setAnchor(new Rectangle2D.Double(0, 0, PAGE_WIDTH_PT, MASTHEAD_HEIGHT_PT));
            masthead.setFillColor(NAVY);
            masthead.setLineColor(null);

            var footer = (XSLFAutoShape) slide.createAutoShape();
            footer.setShapeType(ShapeType.RECT);
            footer.setAnchor(new Rectangle2D.Double(0, PAGE_HEIGHT_PT - 4, PAGE_WIDTH_PT, 4));
            footer.setFillColor(GOLD);
            footer.setLineColor(null);

            addTextBox(slide, new Rectangle2D.Double(32, 10, 300, 34), "SUMMIT & CO", 20, true, Color.WHITE, false);

            if (titleSlide) {
                addTextBox(slide, new Rectangle2D.Double(60, 200, 840, 80), "Team Offsite Budget", 40, true, NAVY, true);
                addTextBox(slide, new Rectangle2D.Double(60, 290, 840, 50), "Q3 2026 Planning Review", 20, false,
                    Color.GRAY, true);
            }

            var out = new ByteArrayOutputStream();
            ppt.write(out);
            return out.toByteArray();
        }
    }

    private static void addTextBox(org.apache.poi.xslf.usermodel.XSLFSlide slide, Rectangle2D anchor, String text,
                                    double fontSize, boolean bold, Color color, boolean centered) {
        var box = (XSLFTextBox) slide.createTextBox();
        box.setAnchor(anchor);
        var para = box.addNewTextParagraph();
        if (centered) {
            para.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.CENTER);
        }
        var run = para.addNewTextRun();
        run.setText(text);
        run.setFontSize(fontSize);
        run.setBold(bold);
        run.setFontColor(color);
    }

    private BudgetOffsiteExample() {}
}
