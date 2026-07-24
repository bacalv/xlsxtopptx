package xlsxtopptx;

import java.util.List;

/**
 * One row to render via {@link SpreadsheetTemplate#render}.
 *
 * <p>{@code type} selects which marker row in the template to clone -- e.g. {@code
 * type("TEMPLATE ROW BOLD")} matches a row whose first column holds {@code %_TEMPLATE ROW BOLD}.
 * {@code data}'s first argument replaces that marker cell's text; the varargs that follow fill
 * the template row's {@code %_0}, {@code %_1}, ... placeholder cells, in order. A type with no
 * placeholders (e.g. a blank spacer row) can skip {@code data} entirely.
 */
public final class SpreadsheetRow {
    private final String type;
    private final String label;
    private final List<Object> values;

    private SpreadsheetRow(String type, String label, List<Object> values) {
        this.type = type;
        this.label = label;
        this.values = values;
    }

    public String type() {
        return type;
    }

    public String label() {
        return label;
    }

    public List<Object> values() {
        return values;
    }

    public static Builder type(String type) {
        return new Builder(type);
    }

    public static final class Builder {
        private final String type;
        private String label = "";
        private List<Object> values = List.of();

        private Builder(String type) {
            this.type = type;
        }

        public Builder data(String label, Object... values) {
            this.label = label;
            this.values = List.of(values);
            return this;
        }

        public SpreadsheetRow build() {
            return new SpreadsheetRow(type, label, values);
        }
    }
}
