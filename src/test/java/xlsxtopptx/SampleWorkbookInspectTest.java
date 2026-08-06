package xlsxtopptx;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;

class SampleWorkbookInspectTest {
    @Test
    void dump() throws Exception {
        IOUtils.setByteArrayMaxOverride(500_000_000);
        var sb = new StringBuilder();
        try (var fis = new FileInputStream("/Users/bac/IdeaProjects/xlsx-to-pptx-java/SampleWorkbook.xlsx");
             var wb = new XSSFWorkbook(fis)) {
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                var sheet = wb.getSheetAt(s);
                sb.append("=== Sheet ").append(s).append(": ").append(sheet.getSheetName())
                    .append(" (last row ").append(sheet.getLastRowNum()).append(") ===\n");
                for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                    var row = sheet.getRow(r);
                    if (row == null) continue;
                    for (int c = 0; c < row.getLastCellNum(); c++) {
                        var cell = row.getCell(c);
                        if (cell == null) continue;
                        var style = (XSSFCellStyle) cell.getCellStyle();
                        var fillPattern = style.getFillPattern();
                        String fillDesc = "none";
                        if (fillPattern == FillPatternType.SOLID_FOREGROUND) {
                            var color = style.getFillForegroundColorColor();
                            if (color != null) {
                                fillDesc = "themed=" + color.isThemed() + " theme=" + color.getTheme()
                                    + " tint=" + color.getTint() + " rgb=" + java.util.Arrays.toString(color.getRGB());
                            }
                        }
                        var type = cell.getCellType();
                        String val = switch (type) {
                            case STRING -> cell.getStringCellValue();
                            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                            case FORMULA -> "=" + cell.getCellFormula();
                            case BLANK -> "";
                            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                            default -> "?" + type;
                        };
                        sb.append(String.format("  [%d,%d] type=%s val=%-20s styleIdx=%d fill=[%s]%n",
                            r, c, type, val, style.getIndex(), fillDesc));
                    }
                }
                sb.append("Merged regions: ").append(sheet.getMergedRegions()).append("\n");
                sb.append("Conditional formattings: ").append(sheet.getSheetConditionalFormatting().getNumConditionalFormattings()).append("\n");
                for (int i = 0; i < sheet.getSheetConditionalFormatting().getNumConditionalFormattings(); i++) {
                    var cf = sheet.getSheetConditionalFormatting().getConditionalFormattingAt(i);
                    sb.append("  CF range=").append(java.util.Arrays.toString(cf.getFormattingRanges())).append("\n");
                }
            }
        }
        Files.writeString(Path.of("/tmp/sample-workbook-dump.txt"), sb.toString());
    }
}
