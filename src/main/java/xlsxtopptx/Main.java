package xlsxtopptx;

import lombok.extern.slf4j.Slf4j;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/** CLI entry point: wires up {@link SpreadsheetToPptx#convert} for command-line use. */
@Slf4j
public final class Main {
    private Main() {}

    public static void main(String[] args) throws IOException {
        if (args.length < 2) {
            log.info("Usage: <input.xlsx> <output.pptx> [sheetIndex]");
            return;
        }
        var sheetIndex = args.length > 2 ? Integer.parseInt(args[2]) : 0;

        byte[] pptxBytes;
        try (var excelInput = new FileInputStream(args[0])) {
            var params = ConvertParams.builder()
                .excelInput(excelInput)
                .sheetIndex(sheetIndex)
                .build();
            pptxBytes = new SpreadsheetToPptx().convert(params);
        }

        try (var out = new FileOutputStream(args[1])) {
            out.write(pptxBytes);
        }
        log.info("Wrote {}", args[1]);
    }
}
