package xlsxtopptx;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class LayoutEngineTest {

    private static CellSnapshot cell(String text) {
        return cell(text, false, LayoutConstants.DEFAULT_FONT_PT);
    }

    private static CellSnapshot cell(String text, boolean bold, double fontSize) {
        return new CellSnapshot(text, null, LayoutConstants.DEFAULT_FONT_COLOR, bold, false,
            fontSize, "Calibri", TextAlign.LEFT, CellBorders.NONE);
    }

    private static SheetSnapshot sheetOf(List<List<CellSnapshot>> grid) {
        var rows = grid.stream().map(RowSnapshot::new).toList();
        var numCols = grid.isEmpty() ? 0 : grid.get(0).size();
        return new SheetSnapshot(rows, List.of(), numCols);
    }

    @Test
    void smallSheetWithAmpleSpaceIsNotShrunk() {
        var sheet = sheetOf(List.of(
            List.of(cell("A"), cell("B")),
            List.of(cell("C"), cell("D"))
        ));

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 1000, 1000);

        assertEquals(1.0, layout.fontScale());
        // wGrow/hGrow are chosen so the sums land on availableWidth/Height exactly in theory, but
        // floating-point arithmetic can overshoot by a hair -- a tiny epsilon avoids a flaky test.
        assertTrue(layout.tableWidth() <= 1000 + 0.01);
        assertTrue(layout.tableHeight() <= 1000 + 0.01);
    }

    @Test
    void wideContentShrinksFontToFitAvailableWidth() {
        var wideText = "X".repeat(200);
        var sheet = sheetOf(List.of(List.of(cell(wideText))));

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 300, 1000);

        assertTrue(layout.fontScale() < 1.0, "expected the long string to force a shrink");
        assertTrue(layout.tableWidth() <= 300 + 0.01, "table must not overflow the available width");
    }

    @Test
    void rowsAreNeverStretchedTallerThanTheirNaturalHeight() {
        // Width is the binding constraint here (one very long cell), so after fontScale is chosen,
        // the sheet's natural height at that scale is well under the generous availableHeight.
        // Rows must NOT be inflated to consume all of it -- that was the "wasted space per row" bug.
        var wideText = "X".repeat(200);
        var sheet = sheetOf(List.of(List.of(cell(wideText))));

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 300, 5000);

        assertTrue(layout.tableHeight() < 5000,
            "a table bound by width should end up shorter than a generously tall slide, not stretched to fill it");
    }

    @Test
    void columnsAreNotStretchedToFillAvailableWidthByDefault() {
        // scaleToFitWidth defaults to off, so a narrow sheet keeps its natural width instead of
        // being stretched to fill the slide -- the caller centers it horizontally instead.
        var sheet = sheetOf(List.of(List.of(cell("hi"), cell("yo"))));

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 1000, 1000);

        assertTrue(layout.tableWidth() < 1000.0);
    }

    @Test
    void scaleToFitWidthStretchesColumnsToFillAvailableWidth() {
        var sheet = sheetOf(List.of(List.of(cell("hi"), cell("yo"))));

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 1000, 1000, true, false);

        assertEquals(1000.0, layout.tableWidth(), 0.01);
    }

    @Test
    void scaleToFitHeightStretchesRowsToFillAvailableHeight() {
        var sheet = sheetOf(List.of(List.of(cell("hi"), cell("yo"))));

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 1000, 1000, false, true);

        assertEquals(1000.0, layout.tableHeight(), 0.01);
    }

    @Test
    void heightBoundSheetShrinksToFitVerticalSpace() {
        var manyRows = java.util.stream.IntStream.range(0, 200)
            .mapToObj(i -> List.of(cell("row" + i)))
            .toList();
        var sheet = sheetOf(manyRows);

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 1000, 300);

        assertTrue(layout.fontScale() < 1.0, "expected 200 rows to force a vertical shrink");
        assertTrue(layout.tableHeight() <= 300 + 0.01);
    }

    @Test
    void extremeSheetsAreClampedAtTheMinimumFontFloorInsteadOfShrinkingToZero() {
        var manyWideRows = java.util.stream.IntStream.range(0, 500)
            .mapToObj(i -> List.of(cell("a very long line of content indeed " + i)))
            .toList();
        var sheet = sheetOf(manyWideRows);

        var layout = LayoutEngine.compute(sheet, MergeIndex.build(List.of()), 50, 50);

        var expectedFloor = LayoutConstants.MIN_FONT_PT / LayoutConstants.DEFAULT_FONT_PT;
        assertEquals(expectedFloor, layout.fontScale(), 1e-9,
            "an impossibly small slide should clamp at the font floor rather than shrinking further");
    }

    @Test
    void aWideMergedAnchorCellDoesNotInflateItsOwnColumnWidth() {
        // The merged cell's long text visually spans two columns, so it shouldn't be measured as
        // if it all had to fit in column 0 alone -- column 0 should end up sized the same as
        // column 1, which is genuinely empty, rather than being driven wide by the long merged text.
        var longText = "a very long banner that spans two columns and then some more padding text";
        var sheet = sheetOf(List.of(List.of(cell(longText), cell(""))));
        var mergeIndex = MergeIndex.build(List.of(new MergeRegion(0, 0, 0, 1)));

        var layout = LayoutEngine.compute(sheet, mergeIndex, 1000, 1000);

        assertEquals(layout.colWidths().get(1), layout.colWidths().get(0), 0.01);
    }
}
