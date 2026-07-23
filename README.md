# xlsx-to-pptx

Converts one Excel sheet into a single PowerPoint slide, scaling font size,
column widths, and row heights uniformly so the whole sheet fits on one
slide. Preserves fill colors, borders, font styling, alignment, and merged
cells (including vertical merges spanning many rows).

## Setup

```
./gradlew build
```

Gradle will pull `org.apache.poi:poi-ooxml` and its transitive deps from
Maven Central on first run.

## Run (CLI)

```
./gradlew run --args="input.xlsx output.pptx"
```

Optional third arg selects a sheet by index (default 0):

```
./gradlew run --args="input.xlsx output.pptx 1"
```

The CLI (`Main`) only exposes `input.xlsx`, `output.pptx`, and an optional
sheet index. For every other option below, call `SpreadsheetToPptx.convert`
directly from Java.

## Use as a library

```java
byte[] pptxBytes;
try (var excelInput = new FileInputStream("input.xlsx")) {
    var params = ConvertParams.builder()
        .excelInput(excelInput)
        .sheetIndex(0)
        .title("Q3 Forecast")
        .scaleToFitWidth(true)
        .build();
    pptxBytes = new SpreadsheetToPptx().convert(params);
}
Files.write(Path.of("output.pptx"), pptxBytes);
```

`convert` takes a single `ConvertParams` (a Lombok `@Builder`) and returns
the generated `.pptx` as a `byte[]` -- it never touches the filesystem
itself, so opening the input and writing the output are the caller's
responsibility.

### `ConvertParams` options

| Option             | Type          | Default                | Description |
|---------------------|---------------|-------------------------|--------------|
| `excelInput`        | `InputStream` | *required*              | The `.xlsx` file to read. Read but never closed by `convert` -- opening and closing it is the caller's job. |
| `sheetIndex`         | `int`         | `0`                      | Which sheet in the workbook to convert. |
| `targetRange`        | `String`      | `null` (whole sheet)     | An Excel A1-notation range, e.g. `"B2:F10"`, restricting the conversion to that rectangle instead of the sheet's full used area. |
| `scaleToFitWidth`    | `boolean`     | `false`                  | Stretch column widths to fill the available width. Left off, a sheet narrower than the target area keeps its natural size and is centered horizontally instead. |
| `scaleToFitHeight`   | `boolean`     | `false`                  | Stretch row heights to fill the available height. Left off, a sheet shorter than the target area keeps its natural size. |
| `templateInput`      | `InputStream` | `null` (blank slide)     | A `.pptx` whose first slide is reused as the base for the output (e.g. one carrying a logo or other branding) instead of creating a blank presentation/slide. Like `excelInput`, read but never closed by `convert`. |
| `tableArea`          | `Rectangle2D` | `null` (default area)    | The exact rectangle (in points, relative to the slide's top-left corner) the table is laid out into, overriding the default margin-based area. Useful together with `templateInput` to fit the table around existing template content. |
| `title`              | `String`      | `null` (no title)        | Rendered as a bold title textbox at the top of the slide. Left unset, no title is added at all. |

## Testing

```
./gradlew test
```

## Background

This started as a direct Java/Apache POI port of a Python prototype
(openpyxl + python-pptx) that validated the core algorithm -- real merges,
no text wrapping, uniform font-scale-to-fit, graceful degradation on a
large stress-test sheet -- by rendering to PNG via
`soffice --headless --convert-to png` and checking visually. The POI port
has since grown its own test suite (`./gradlew test`) covering layout,
merges, borders, and the CLI's argument handling.
