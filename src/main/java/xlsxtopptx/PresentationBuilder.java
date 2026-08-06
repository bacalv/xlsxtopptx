package xlsxtopptx;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Merges a multi-slide .pptx template with one .xlsx workbook, resolving two kinds of markers
 * anywhere in the template's text:
 *
 * <p>Any shape whose own text starts with {@code %_SHEET(SheetName)} or
 * {@code %_SHEET(SheetName!B2:F10)} is deleted and replaced by a table rendered from that sheet
 * (the whole used range, or the given range), sized and positioned to exactly fill the shape's own
 * bounding box -- stretched to fit both width and height. A shape whose text merely contains
 * {@code %_SHEET(...)} somewhere other than the start is left untouched.
 *
 * <p>Separately, {@code $NAME} tokens anywhere else in the deck's text are replaced by values
 * registered via {@link #placeholder}, mirroring {@link SpreadsheetTemplate}'s identical mechanism
 * for Excel cells: every token found in the deck must have a matching {@link #placeholder} call
 * and vice versa, checked once for the whole deck before any substitution happens.
 *
 * <p>Both failure modes -- a malformed/unresolvable {@code %_SHEET(...)} directive, or a
 * {@code $NAME} mismatch -- throw {@link IllegalArgumentException} rather than skipping the
 * offending shape, since either almost always indicates a typo or a stale template.
 *
 * <p>Scope, kept simple on purpose: only top-level shapes on each slide are scanned -- shapes
 * nested inside a grouped shape are not.
 */
public final class PresentationBuilder {
    private static final Pattern SHEET_DIRECTIVE =
        Pattern.compile("^%_SHEET\\(\\s*([^!()]+?)\\s*(?:!\\s*([^()]+?)\\s*)?\\)");
    private static final Pattern PLACEHOLDER_TOKEN = Pattern.compile("\\$([A-Za-z_][A-Za-z0-9_]*)");

    private final XMLSlideShow ppt;
    private final XSSFWorkbook workbook;
    private final Map<String, String> placeholders = new HashMap<>();

    private PresentationBuilder(XMLSlideShow ppt, XSSFWorkbook workbook) {
        this.ppt = ppt;
        this.workbook = workbook;
    }

    /** {@code templateInput} and {@code excelInput} are both fully read here; closing them
     *  afterward is the caller's responsibility but is a no-op since POI already reads each stream
     *  to completion during construction. */
    public static PresentationBuilder of(InputStream templateInput, InputStream excelInput) throws IOException {
        return new PresentationBuilder(new XMLSlideShow(templateInput), new XSSFWorkbook(excelInput));
    }

    /** Registers a deck-wide {@code $name} token (case as given, e.g. {@code "REPORT_DATE"} for
     *  {@code $REPORT_DATE}) to be replaced by {@code value} everywhere it appears in the
     *  template's text. Must be called before {@link #build}. */
    public PresentationBuilder placeholder(String name, Object value) {
        placeholders.put(name, String.valueOf(value));
        return this;
    }

    public byte[] build() throws IOException {
        try (var slideshow = ppt; var wb = workbook) {
            replaceSheetDirectives();
            substitutePlaceholders();

            var out = new ByteArrayOutputStream();
            slideshow.write(out);
            return out.toByteArray();
        }
    }

    // ---------- %_SHEET(...) ----------

    private void replaceSheetDirectives() {
        for (int slideIdx = 0; slideIdx < ppt.getSlides().size(); slideIdx++) {
            var slide = ppt.getSlides().get(slideIdx);
            // Snapshot the shape list before mutating it -- removeShape below would otherwise
            // invalidate the live list this loop is iterating.
            for (var shape : List.copyOf(slide.getShapes())) {
                if (!(shape instanceof XSLFTextShape textShape)) continue;

                var text = textShape.getText();
                var trimmed = text == null ? "" : text.stripLeading();
                if (!trimmed.startsWith("%_SHEET(")) continue;

                applySheetDirective(slide, slideIdx, textShape, trimmed);
            }
        }
    }

    private void applySheetDirective(XSLFSlide slide, int slideIdx, XSLFTextShape shape, String directiveText) {
        var matcher = SHEET_DIRECTIVE.matcher(directiveText);
        if (!matcher.lookingAt()) {
            throw new IllegalArgumentException(
                "Malformed %_SHEET(...) directive on slide " + (slideIdx + 1) + ": \"" + directiveText + "\"");
        }

        var sheetName = matcher.group(1).trim();
        var rangeText = matcher.group(2);

        var sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Unknown sheet \"" + sheetName + "\" in %_SHEET(...) directive on slide "
                + (slideIdx + 1) + " -- workbook has: " + sheetNames());
        }

        var bounds = resolveDirectiveBounds(sheet, rangeText, slideIdx);

        SpreadsheetToPptx.layoutTable(slide, sheet, bounds, shape.getAnchor(), true, true);
        slide.removeShape(shape);
    }

    private CellRangeAddress resolveDirectiveBounds(Sheet sheet, String rangeText, int slideIdx) {
        if (rangeText == null) {
            return SheetReader.usedBounds(sheet);
        }
        try {
            return CellRangeAddress.valueOf(rangeText.trim());
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("Invalid range \"" + rangeText + "\" in %_SHEET(...) directive on slide "
                + (slideIdx + 1) + ": " + e.getMessage(), e);
        }
    }

    private List<String> sheetNames() {
        var names = new ArrayList<String>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            names.add(workbook.getSheetName(i));
        }
        return names;
    }

    // ---------- $NAME ----------

    /** Replaces every {@code $NAME} token in the deck's remaining shapes (i.e. after
     *  {@link #replaceSheetDirectives} has removed every {@code %_SHEET(...)} shape, so generated
     *  table cells are never scanned) with its registered {@link #placeholder} value, after
     *  checking that the tokens actually present exactly match what was registered. */
    private void substitutePlaceholders() {
        var foundTokens = new HashSet<String>();
        var matchingParagraphs = new ArrayList<XSLFTextParagraph>();

        for (var slide : ppt.getSlides()) {
            for (var shape : slide.getShapes()) {
                if (!(shape instanceof XSLFTextShape textShape)) continue;
                for (var para : textShape.getTextParagraphs()) {
                    var paraText = paragraphText(para);
                    if (PLACEHOLDER_TOKEN.matcher(paraText).find()) {
                        collectTokens(paraText, foundTokens);
                        matchingParagraphs.add(para);
                    }
                }
            }
        }

        var unknownTokens = foundTokens.stream().filter(t -> !placeholders.containsKey(t)).toList();
        if (!unknownTokens.isEmpty()) {
            throw new IllegalArgumentException(
                "Unknown placeholder(s) in template: " + unknownTokens + " -- registered placeholder(s): "
                    + placeholders.keySet());
        }
        var unusedPlaceholders = placeholders.keySet().stream().filter(p -> !foundTokens.contains(p)).toList();
        if (!unusedPlaceholders.isEmpty()) {
            throw new IllegalArgumentException(
                "Placeholder(s) registered but not found in template: " + unusedPlaceholders);
        }

        for (var para : matchingParagraphs) {
            substituteInParagraph(para);
        }
    }

    private static String paragraphText(XSLFTextParagraph para) {
        var sb = new StringBuilder();
        for (var run : para.getTextRuns()) {
            sb.append(run.getRawText());
        }
        return sb.toString();
    }

    private static void collectTokens(String text, Set<String> into) {
        var matcher = PLACEHOLDER_TOKEN.matcher(text);
        while (matcher.find()) {
            into.add(matcher.group(1));
        }
    }

    /** A {@code $TOKEN} match's position within a paragraph, expressed relative to whichever
     *  single run it falls entirely inside. */
    private record TokenMatch(int runIndex, int localStart, int localEnd, String replacement) {}

    /** Replaces every {@code $TOKEN} in {@code para} with its registered value, run by run, so
     *  that each run's own formatting (bold, color, ...) is preserved untouched. A token whose
     *  characters straddle two runs (e.g. part of it re-styled in the authoring tool) has no
     *  well-defined "which run's formatting wins", so that's a hard failure rather than a guess. */
    private void substituteInParagraph(XSLFTextParagraph para) {
        var runs = para.getTextRuns();
        var runTexts = runs.stream().map(XSLFTextRun::getRawText).toList();

        var runStarts = new int[runs.size()];
        var offset = 0;
        for (int i = 0; i < runs.size(); i++) {
            runStarts[i] = offset;
            offset += runTexts.get(i).length();
        }

        var fullText = String.join("", runTexts);
        var matcher = PLACEHOLDER_TOKEN.matcher(fullText);
        var matches = new ArrayList<TokenMatch>();
        while (matcher.find()) {
            var start = matcher.start();
            var end = matcher.end();
            var startRun = runIndexAt(runStarts, start);
            var endRun = runIndexAt(runStarts, end - 1);
            if (startRun != endRun) {
                throw new IllegalArgumentException("Placeholder $" + matcher.group(1)
                    + " is split across multiple differently-formatted text runs -- retype it as a single run");
            }
            matches.add(new TokenMatch(startRun, start - runStarts[startRun], end - runStarts[startRun],
                placeholders.get(matcher.group(1))));
        }

        // Applied back-to-front so a replacement's length change never invalidates an
        // earlier-in-the-run match's still-pending local offsets.
        for (int i = matches.size() - 1; i >= 0; i--) {
            var m = matches.get(i);
            var run = runs.get(m.runIndex());
            var text = run.getRawText();
            run.setText(text.substring(0, m.localStart()) + m.replacement() + text.substring(m.localEnd()));
        }
    }

    private static int runIndexAt(int[] runStarts, int globalOffset) {
        for (int i = runStarts.length - 1; i >= 0; i--) {
            if (globalOffset >= runStarts[i]) return i;
        }
        return 0;
    }
}
