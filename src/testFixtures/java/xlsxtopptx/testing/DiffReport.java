package xlsxtopptx.testing;

import java.util.List;

/** Turns a list of human-readable difference lines into a single {@link AssertionError}, capped
 *  so a wildly different pair of files doesn't dump thousands of lines into a test report. */
final class DiffReport {
    private static final int MAX_DIFFS_SHOWN = 50;

    private DiffReport() {}

    static void failIfAny(String userMessage, String subject, List<String> diffs) {
        if (diffs.isEmpty()) return;

        var sb = new StringBuilder();
        if (userMessage != null && !userMessage.isBlank()) {
            sb.append(userMessage).append(System.lineSeparator());
        }
        sb.append(subject).append(" - ").append(diffs.size())
            .append(diffs.size() == 1 ? " difference" : " differences").append(":");

        var shown = diffs.subList(0, Math.min(diffs.size(), MAX_DIFFS_SHOWN));
        for (var diff : shown) {
            sb.append(System.lineSeparator()).append("  - ").append(diff);
        }
        if (diffs.size() > shown.size()) {
            sb.append(System.lineSeparator()).append("  ... and ").append(diffs.size() - shown.size()).append(" more");
        }

        throw new AssertionError(sb.toString());
    }
}
