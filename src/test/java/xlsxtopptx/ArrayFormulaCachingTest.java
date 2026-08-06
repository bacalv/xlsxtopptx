package xlsxtopptx;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;

class ArrayFormulaCachingTest {
    @Test
    void beforeAndAfterEvaluatingAnchor() throws Exception {
        var sb = new StringBuilder();
        try (var fis = new FileInputStream("/Users/bac/IdeaProjects/xlsx-to-pptx-java/SampleWorkbook.xlsx");
             var wb = new XSSFWorkbook(fis)) {
            var sheet = wb.getSheet("OUTPUT_2");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();

            dumpRaw(sb, sheet, "BEFORE any evaluation");

            var a1 = sheet.getRow(0).getCell(0);
            sb.append("A1 type=").append(a1.getCellType()).append(" formula=")
                .append(a1.getCellType() == CellType.FORMULA ? a1.getCellFormula() : "n/a").append("\n");
            var evaluated = evaluator.evaluate(a1);
            sb.append("evaluate(A1) result=").append(evaluated).append("\n");

            dumpRaw(sb, sheet, "AFTER evaluating A1");
        }
        Files.writeString(Path.of("/tmp/sample-workbook-array-formula-dump.txt"), sb.toString());
    }

    private void dumpRaw(StringBuilder sb, org.apache.poi.ss.usermodel.Sheet sheet, String label) {
        sb.append("=== ").append(label).append(" ===\n");
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            var row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                var cell = row.getCell(c);
                if (cell == null) continue;
                var type = cell.getCellType();
                String val = switch (type) {
                    case STRING -> cell.getStringCellValue();
                    case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                    case FORMULA -> "=" + cell.getCellFormula() + " (cached=" + cachedStr(cell) + ")";
                    case BLANK -> "<blank>";
                    default -> "?" + type;
                };
                sb.append(String.format("  [%d,%d] type=%s val=%s%n", r, c, type, val));
            }
        }
    }

    private String cachedStr(org.apache.poi.ss.usermodel.Cell cell) {
        try {
            return switch (cell.getCachedFormulaResultType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                default -> "?" + cell.getCachedFormulaResultType();
            };
        } catch (Exception e) {
            return "err:" + e.getMessage();
        }
    }
}
