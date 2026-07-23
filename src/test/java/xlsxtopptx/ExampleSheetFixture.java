package xlsxtopptx;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Builds a small but structurally faithful stand-in for the real "Funding Requests" sheet this
 * converter was built and tested against: a two-level merged banner header, bold sub-headers, a
 * bold summary row carrying a light-blue bottom border, a blank spacer row, plain data rows, and
 * a navy/white bold total row. Six columns and eight rows instead of fourteen and thirty, but
 * every layout feature the converter has to handle is represented.
 */
final class ExampleSheetFixture {
    private ExampleSheetFixture() {}

    static final String SHEET_NAME = "Funding Requests";
    static final byte[] NAVY = {(byte) 0x1F, (byte) 0x1F, (byte) 0x6E};
    static final byte[] LIGHT_BLUE = {(byte) 0xAD, (byte) 0xD8, (byte) 0xE6};
    static final String NUMBER_FORMAT = "#,##0.00";

    // Row indices.
    static final int ROW_YEAR_BANNER = 0;
    static final int ROW_REQUEST_BANNER = 1;
    static final int ROW_SUB_HEADERS = 2;
    static final int ROW_SUMMARY = 3;
    static final int ROW_SPACER = 4;
    static final int ROW_DATA_1 = 5;
    static final int ROW_DATA_2 = 6;
    static final int ROW_TOTAL = 7;
    static final int ROW_COUNT = 8;

    // Column indices.
    static final int COL_LABEL = 0;
    static final int COL_FIRST_METRIC = 1;
    static final int COL_LAST_METRIC = 4;
    static final int COL_TOTAL = 5;
    static final int COL_COUNT = 6;

    static XSSFWorkbook build() {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.createSheet(SHEET_NAME);

        var boldLabel = boldStyle(workbook);
        var navyBanner = navyBannerStyle(workbook);
        var subHeader = boldCenteredStyle(workbook);
        var summaryNumber = summaryNumberStyle(workbook);
        var plainNumber = plainNumberStyle(workbook);
        var navyLabel = navyStyle(workbook, false);
        var navyNumber = navyStyle(workbook, true);

        var yearBannerRow = sheet.createRow(ROW_YEAR_BANNER);
        setCell(yearBannerRow, 2, "FY 2026 Forecast", navyBanner);
        sheet.addMergedRegion(new CellRangeAddress(ROW_YEAR_BANNER, ROW_YEAR_BANNER, 2, 3));

        var requestBannerRow = sheet.createRow(ROW_REQUEST_BANNER);
        setCell(requestBannerRow, COL_LABEL, "Technology [L4] as Sponsor", boldLabel);
        setCell(requestBannerRow, 1, "FY'26 Funding Requests", navyBanner);
        sheet.addMergedRegion(new CellRangeAddress(ROW_REQUEST_BANNER, ROW_REQUEST_BANNER, 1, COL_TOTAL));

        var subHeaderRow = sheet.createRow(ROW_SUB_HEADERS);
        setCell(subHeaderRow, 1, "PS/BS", subHeader);
        setCell(subHeaderRow, 2, "CO", subHeader);
        setCell(subHeaderRow, 3, "Bus Mand", subHeader);
        setCell(subHeaderRow, 4, "Tech Mand", subHeader);
        setCell(subHeaderRow, COL_TOTAL, "Total", subHeader);

        var summaryRow = sheet.createRow(ROW_SUMMARY);
        setCell(summaryRow, COL_LABEL, "Products & Functions Technology [L5]", boldLabel);
        setNumericRow(summaryRow, summaryNumber, 2145, 45, 156, 11223, 13581);

        sheet.createRow(ROW_SPACER); // deliberately blank

        var data1 = sheet.createRow(ROW_DATA_1);
        setCell(data1, COL_LABEL, "Item One", null);
        setNumericRow(data1, plainNumber, 10, 5, 3, 2, 20);

        var data2 = sheet.createRow(ROW_DATA_2);
        setCell(data2, COL_LABEL, "Item Two", null);
        setNumericRow(data2, plainNumber, 50, 25, 8, 1, 84);

        var totalRow = sheet.createRow(ROW_TOTAL);
        setCell(totalRow, COL_LABEL, "Total Funding Requests", navyLabel);
        setNumericRow(totalRow, navyNumber, 60, 30, 11, 3, 104);

        for (int c = 0; c < COL_COUNT; c++) {
            sheet.setColumnWidth(c, 20 * 256);
        }

        return workbook;
    }

    private static void setNumericRow(org.apache.poi.ss.usermodel.Row row, XSSFCellStyle style, double... values) {
        for (int i = 0; i < values.length; i++) {
            var cell = row.createCell(COL_FIRST_METRIC + i);
            cell.setCellValue(values[i]);
            cell.setCellStyle(style);
        }
    }

    private static void setCell(org.apache.poi.ss.usermodel.Row row, int col, String text, XSSFCellStyle style) {
        var cell = row.createCell(col);
        cell.setCellValue(text);
        if (style != null) cell.setCellStyle(style);
    }

    private static XSSFCellStyle boldStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook, false));
        return style;
    }

    private static XSSFCellStyle boldCenteredStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook, false));
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private static XSSFCellStyle navyBannerStyle(XSSFWorkbook workbook) {
        var style = navyStyle(workbook, false);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private static XSSFCellStyle navyStyle(XSSFWorkbook workbook, boolean numeric) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook, true));
        style.setFillForegroundColor(new XSSFColor(NAVY));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        if (numeric) {
            style.setDataFormat(workbook.createDataFormat().getFormat(NUMBER_FORMAT));
            style.setAlignment(HorizontalAlignment.RIGHT);
        }
        return style;
    }

    private static XSSFCellStyle summaryNumberStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setFont(boldFont(workbook, false));
        style.setDataFormat(workbook.createDataFormat().getFormat(NUMBER_FORMAT));
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(new XSSFColor(LIGHT_BLUE));
        // Simulate a real-world writer (e.g. openpyxl) that defines a genuine border but never
        // emits applyBorder="1" -- SheetReader must still pick this border up.
        style.getCoreXf().unsetApplyBorder();
        return style;
    }

    private static XSSFCellStyle plainNumberStyle(XSSFWorkbook workbook) {
        var style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(NUMBER_FORMAT));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    private static XSSFFont boldFont(XSSFWorkbook workbook, boolean white) {
        var font = workbook.createFont();
        font.setBold(true);
        if (white) font.setColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF}));
        return font;
    }
}
