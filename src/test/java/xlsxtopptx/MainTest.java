package xlsxtopptx;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Covers {@link Main#main}'s command-line argument handling -- the usage message for too few
 * arguments, sheetIndex parsing, and failure modes for bad paths/indices -- separate from
 * {@link SpreadsheetToPptxAcceptanceTest}, which exercises {@link SpreadsheetToPptx#convert}
 * directly.
 */
class MainTest {

    @TempDir
    Path tempDir;

    private Path xlsxPath;
    private Path pptxPath;

    @BeforeEach
    void generateFixtureWorkbook() throws Exception {
        xlsxPath = tempDir.resolve("example.xlsx");
        pptxPath = tempDir.resolve("example.pptx");

        try (var workbook = ExampleSheetFixture.build();
             var out = new FileOutputStream(xlsxPath.toFile())) {
            workbook.write(out);
        }
    }

    @Test
    void noArgumentsPrintsUsageAndWritesNothing() {
        assertDoesNotThrow(() -> Main.main(new String[0]));

        assertFalse(Files.exists(pptxPath));
    }

    @Test
    void aSingleArgumentPrintsUsageAndWritesNothing() {
        var args = new String[]{xlsxPath.toString()};
        assertDoesNotThrow(() -> Main.main(args));

        assertFalse(Files.exists(pptxPath));
    }

    @Test
    void validArgumentsConvertTheDefaultSheet() throws Exception {
        Main.main(new String[]{xlsxPath.toString(), pptxPath.toString()});

        assertTrue(Files.exists(pptxPath));
        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            assertEquals(1, ppt.getSlides().size());
        }
    }

    @Test
    void explicitSheetIndexIsParsedAndUsed() throws Exception {
        Main.main(new String[]{xlsxPath.toString(), pptxPath.toString(), "0"});

        assertTrue(Files.exists(pptxPath));
        try (var ppt = new XMLSlideShow(Files.newInputStream(pptxPath))) {
            assertEquals(1, ppt.getSlides().size());
        }
    }

    @Test
    void nonNumericSheetIndexFailsFast() {
        var args = new String[]{xlsxPath.toString(), pptxPath.toString(), "not-a-number"};

        assertThrows(NumberFormatException.class, () -> Main.main(args));

        assertFalse(Files.exists(pptxPath), "no output file should be written when argument parsing fails");
    }

    @Test
    void outOfRangeSheetIndexFailsFast() {
        var args = new String[]{xlsxPath.toString(), pptxPath.toString(), "99"};

        assertThrows(IllegalArgumentException.class, () -> Main.main(args));

        assertFalse(Files.exists(pptxPath), "no output file should be written when the sheet index is invalid");
    }

    @Test
    void missingInputFileFailsFast() {
        var missing = tempDir.resolve("does-not-exist.xlsx");
        var args = new String[]{missing.toString(), pptxPath.toString()};

        assertThrows(FileNotFoundException.class, () -> Main.main(args));

        assertFalse(Files.exists(pptxPath), "no output file should be written when the input file is missing");
    }
}
