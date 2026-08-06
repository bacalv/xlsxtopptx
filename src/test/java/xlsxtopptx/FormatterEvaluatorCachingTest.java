package xlsxtopptx;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;

class FormatterEvaluatorCachingTest {
    @Test
    void formatEachCellFreshEvaluatorEachTime() throws Exception {
        var sb = new StringBuilder();
        try (var fis = new FileInputStream("/Users/bac/IdeaProjects/xlsx-to-pptx-java/SampleWorkbook.xlsx");
             var wb = new XSSFWorkbook(fis)) {
            var sheet = wb.getSheet("OUTPUT_2");
            var formatter = new DataFormatter();

            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                var row = sheet.getRow(r);
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    var cell = row.getCell(c);
                    // fresh evaluator per cell, like giving each its own clean-slate evaluate call
                    var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                    var text = formatter.formatCellValue(cell, evaluator);
                    sb.append(String.format("[%d,%d] freshEval=%s%n", r, c, text));
                }
            }

            sb.append("--- now with ONE shared evaluator reused across all cells (like SheetReader does) ---\n");
            var sharedEvaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                var row = sheet.getRow(r);
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    var cell = row.getCell(c);
                    var text = formatter.formatCellValue(cell, sharedEvaluator);
                    sb.append(String.format("[%d,%d] sharedEval=%s%n", r, c, text));
                }
            }
        }
        Files.writeString(Path.of("/tmp/sample-workbook-formatter-dump.txt"), sb.toString());
    }
}
