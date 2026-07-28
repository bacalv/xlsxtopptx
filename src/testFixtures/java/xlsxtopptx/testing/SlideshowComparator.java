package xlsxtopptx.testing;

import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Compares two slideshows shape-by-shape (position/size, text/run styling, table cells, picture
 * bytes), ignoring the incidental byte-level noise between two writer runs of the "same"
 * presentation: core.xml timestamps, relationship id numbering, auto-generated shape names, and
 * XML formatting/ordering.
 */
final class SlideshowComparator {
    private static final double TOLERANCE_PT = 0.5;

    private SlideshowComparator() {}

    static List<String> compare(XMLSlideShow expected, XMLSlideShow actual) {
        var diffs = new ArrayList<String>();

        if (!expected.getPageSize().equals(actual.getPageSize())) {
            diffs.add("page size differs: expected=%s actual=%s".formatted(expected.getPageSize(), actual.getPageSize()));
        }

        var expectedSlides = expected.getSlides();
        var actualSlides = actual.getSlides();
        if (expectedSlides.size() != actualSlides.size()) {
            diffs.add("slide count differs: expected=%d actual=%d".formatted(expectedSlides.size(), actualSlides.size()));
        }

        var n = Math.min(expectedSlides.size(), actualSlides.size());
        for (int i = 0; i < n; i++) {
            compareSlide("slide[%d]".formatted(i), expectedSlides.get(i), actualSlides.get(i), diffs);
        }
        return diffs;
    }

    private static void compareSlide(String prefix, XSLFSlide expected, XSLFSlide actual, List<String> diffs) {
        var expectedBg = solidBackgroundColor(expected);
        var actualBg = solidBackgroundColor(actual);
        if (!Objects.equals(expectedBg, actualBg)) {
            diffs.add("%s: background color differs: expected=%s actual=%s".formatted(prefix, colorStr(expectedBg), colorStr(actualBg)));
        }
        compareShapes(prefix, expected.getShapes(), actual.getShapes(), diffs);
    }

    private static Color solidBackgroundColor(XSLFSlide slide) {
        var background = slide.getBackground();
        return background == null ? null : background.getFillColor();
    }

    private static void compareShapes(String prefix, List<XSLFShape> expected, List<XSLFShape> actual, List<String> diffs) {
        if (expected.size() != actual.size()) {
            diffs.add("%s: shape count differs: expected=%d actual=%d".formatted(prefix, expected.size(), actual.size()));
        }
        var n = Math.min(expected.size(), actual.size());
        for (int i = 0; i < n; i++) {
            compareShape("%s.shape[%d]".formatted(prefix, i), expected.get(i), actual.get(i), diffs);
        }
    }

    private static void compareShape(String prefix, XSLFShape expected, XSLFShape actual, List<String> diffs) {
        if (!expected.getClass().equals(actual.getClass())) {
            diffs.add("%s: shape type differs: expected=%s actual=%s".formatted(
                prefix, expected.getClass().getSimpleName(), actual.getClass().getSimpleName()));
            return;
        }

        compareAnchor(prefix, expected.getAnchor(), actual.getAnchor(), diffs);

        // XSLFTableCell and XSLFGroupShape sit above XSLFTextShape/XSLFShape in the hierarchy, so
        // the more specific checks have to run first.
        if (expected instanceof XSLFTable expectedTable && actual instanceof XSLFTable actualTable) {
            compareTable(prefix, expectedTable, actualTable, diffs);
        } else if (expected instanceof XSLFPictureShape expectedPic && actual instanceof XSLFPictureShape actualPic) {
            comparePicture(prefix, expectedPic, actualPic, diffs);
        } else if (expected instanceof XSLFGroupShape expectedGroup && actual instanceof XSLFGroupShape actualGroup) {
            compareShapes(prefix, expectedGroup.getShapes(), actualGroup.getShapes(), diffs);
        } else if (expected instanceof XSLFTextShape expectedText && actual instanceof XSLFTextShape actualText) {
            compareText(prefix, expectedText, actualText, diffs);
        }
    }

    private static void compareAnchor(String prefix, Rectangle2D expected, Rectangle2D actual, List<String> diffs) {
        if (Math.abs(expected.getX() - actual.getX()) > TOLERANCE_PT
            || Math.abs(expected.getY() - actual.getY()) > TOLERANCE_PT
            || Math.abs(expected.getWidth() - actual.getWidth()) > TOLERANCE_PT
            || Math.abs(expected.getHeight() - actual.getHeight()) > TOLERANCE_PT) {
            diffs.add("%s: anchor differs: expected=%s actual=%s".formatted(prefix, rectStr(expected), rectStr(actual)));
        }
    }

    private static void compareText(String prefix, XSLFTextShape expected, XSLFTextShape actual, List<String> diffs) {
        if (!Objects.equals(expected.getText(), actual.getText())) {
            diffs.add("%s: text differs: expected=\"%s\" actual=\"%s\"".formatted(prefix, expected.getText(), actual.getText()));
        }
        compareRuns(prefix, flattenRuns(expected), flattenRuns(actual), diffs);
    }

    private static List<XSLFTextRun> flattenRuns(XSLFTextShape shape) {
        var runs = new ArrayList<XSLFTextRun>();
        for (var paragraph : shape.getTextParagraphs()) {
            runs.addAll(paragraph.getTextRuns());
        }
        return runs;
    }

    private static void compareRuns(String prefix, List<XSLFTextRun> expected, List<XSLFTextRun> actual, List<String> diffs) {
        if (expected.size() != actual.size()) {
            diffs.add("%s: text run count differs: expected=%d actual=%d".formatted(prefix, expected.size(), actual.size()));
            return;
        }
        for (int i = 0; i < expected.size(); i++) {
            compareRun("%s.run[%d]".formatted(prefix, i), expected.get(i), actual.get(i), diffs);
        }
    }

    private static void compareRun(String prefix, XSLFTextRun expected, XSLFTextRun actual, List<String> diffs) {
        var fieldDiffs = new ArrayList<String>();
        if (!Objects.equals(expected.getRawText(), actual.getRawText())) {
            fieldDiffs.add("text: expected=\"%s\" actual=\"%s\"".formatted(expected.getRawText(), actual.getRawText()));
        }
        if (expected.isBold() != actual.isBold()) {
            fieldDiffs.add("bold: expected=%s actual=%s".formatted(expected.isBold(), actual.isBold()));
        }
        if (expected.isItalic() != actual.isItalic()) {
            fieldDiffs.add("italic: expected=%s actual=%s".formatted(expected.isItalic(), actual.isItalic()));
        }
        if (!Objects.equals(expected.getFontSize(), actual.getFontSize())) {
            fieldDiffs.add("fontSize: expected=%s actual=%s".formatted(expected.getFontSize(), actual.getFontSize()));
        }
        var expectedColor = solidColor(expected.getFontColor());
        var actualColor = solidColor(actual.getFontColor());
        if (!Objects.equals(expectedColor, actualColor)) {
            fieldDiffs.add("fontColor: expected=%s actual=%s".formatted(colorStr(expectedColor), colorStr(actualColor)));
        }
        var expectedLink = expected.getHyperlink() != null ? expected.getHyperlink().getAddress() : null;
        var actualLink = actual.getHyperlink() != null ? actual.getHyperlink().getAddress() : null;
        if (!Objects.equals(expectedLink, actualLink)) {
            fieldDiffs.add("hyperlink: expected=%s actual=%s".formatted(expectedLink, actualLink));
        }

        if (!fieldDiffs.isEmpty()) {
            diffs.add("%s: %s".formatted(prefix, String.join(", ", fieldDiffs)));
        }
    }

    private static Color solidColor(PaintStyle paint) {
        return paint instanceof PaintStyle.SolidPaint solid ? solid.getSolidColor().getColor() : null;
    }

    private static void compareTable(String prefix, XSLFTable expected, XSLFTable actual, List<String> diffs) {
        if (expected.getNumberOfRows() != actual.getNumberOfRows() || expected.getNumberOfColumns() != actual.getNumberOfColumns()) {
            diffs.add("%s: table dimensions differ: expected=%dx%d actual=%dx%d".formatted(
                prefix, expected.getNumberOfRows(), expected.getNumberOfColumns(),
                actual.getNumberOfRows(), actual.getNumberOfColumns()));
            return;
        }
        for (int r = 0; r < expected.getNumberOfRows(); r++) {
            for (int c = 0; c < expected.getNumberOfColumns(); c++) {
                compareTableCell("%s.cell[%d,%d]".formatted(prefix, r, c), expected.getCell(r, c), actual.getCell(r, c), diffs);
            }
        }
    }

    private static void compareTableCell(String prefix, XSLFTableCell expected, XSLFTableCell actual, List<String> diffs) {
        var fieldDiffs = new ArrayList<String>();
        if (!Objects.equals(expected.getText(), actual.getText())) {
            fieldDiffs.add("text: expected=\"%s\" actual=\"%s\"".formatted(expected.getText(), actual.getText()));
        }
        if (expected.getGridSpan() != actual.getGridSpan()) {
            fieldDiffs.add("gridSpan: expected=%d actual=%d".formatted(expected.getGridSpan(), actual.getGridSpan()));
        }
        if (expected.getRowSpan() != actual.getRowSpan()) {
            fieldDiffs.add("rowSpan: expected=%d actual=%d".formatted(expected.getRowSpan(), actual.getRowSpan()));
        }
        if (expected.isMerged() != actual.isMerged()) {
            fieldDiffs.add("merged: expected=%s actual=%s".formatted(expected.isMerged(), actual.isMerged()));
        }
        if (!Objects.equals(expected.getFillColor(), actual.getFillColor())) {
            fieldDiffs.add("fill: expected=%s actual=%s".formatted(colorStr(expected.getFillColor()), colorStr(actual.getFillColor())));
        }
        for (var edge : BorderEdge.values()) {
            var expectedColor = expected.getBorderColor(edge);
            var actualColor = actual.getBorderColor(edge);
            if (!Objects.equals(expectedColor, actualColor)) {
                fieldDiffs.add("border[%s]: expected=%s actual=%s".formatted(edge, colorStr(expectedColor), colorStr(actualColor)));
            }
        }

        if (!fieldDiffs.isEmpty()) {
            diffs.add("%s: %s".formatted(prefix, String.join(", ", fieldDiffs)));
        }
        compareRuns(prefix, flattenRuns(expected), flattenRuns(actual), diffs);
    }

    private static void comparePicture(String prefix, XSLFPictureShape expected, XSLFPictureShape actual, List<String> diffs) {
        var expectedData = expected.getPictureData();
        var actualData = actual.getPictureData();
        if (expectedData == null || actualData == null) {
            if (expectedData != actualData) {
                diffs.add("%s: picture data presence differs: expected=%s actual=%s".formatted(
                    prefix, expectedData != null, actualData != null));
            }
            return;
        }
        if (!Arrays.equals(expectedData.getData(), actualData.getData())) {
            diffs.add("%s: picture bytes differ".formatted(prefix));
        }
    }

    private static String colorStr(Color c) {
        return c == null ? "null" : "#%02X%02X%02X".formatted(c.getRed(), c.getGreen(), c.getBlue());
    }

    private static String rectStr(Rectangle2D r) {
        return "[x=%.1f, y=%.1f, w=%.1f, h=%.1f]".formatted(r.getX(), r.getY(), r.getWidth(), r.getHeight());
    }
}
