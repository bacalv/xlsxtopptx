package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

class SampleWorkbookConvertTest {
    @Test
    void convertOutput1() throws Exception {
        convertSheet(1, "/tmp/output1.pptx", "/tmp/output1-dump.txt");
    }

    @Test
    void convertOutput2() throws Exception {
        convertSheet(2, "/tmp/output2.pptx", "/tmp/output2-dump.txt");
    }

    private void convertSheet(int sheetIndex, String outPptx, String dumpPath) throws Exception {
        byte[] pptxBytes;
        try (var excelInput = new FileInputStream("/Users/bac/IdeaProjects/xlsx-to-pptx-java/SampleWorkbook.xlsx")) {
            var params = ConvertParams.builder()
                .excelInput(excelInput)
                .sheetIndex(sheetIndex)
                .build();
            pptxBytes = new SpreadsheetToPptx().convert(params);
        }
        try (var out = new FileOutputStream(outPptx)) {
            out.write(pptxBytes);
        }

        var sb = new StringBuilder();
        try (var ppt = new XMLSlideShow(new java.io.ByteArrayInputStream(pptxBytes))) {
            var slide = ppt.getSlides().get(0);
            for (var shape : slide.getShapes()) {
                if (shape instanceof org.apache.poi.xslf.usermodel.XSLFTable table) {
                    for (var row : table.getRows()) {
                        for (var cell : row.getCells()) {
                            var fill = cell.getFillColor();
                            var text = cell.getText();
                            sb.append(String.format("text=%-20s fill=%s%n", text, fill));
                        }
                        sb.append("---\n");
                    }
                }
            }
        }
        Files.writeString(Path.of(dumpPath), sb.toString());
    }
}
