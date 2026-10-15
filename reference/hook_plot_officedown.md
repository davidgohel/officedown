# knitr hook for figure caption autonumbering

Plot hook for Word documents that generates auto-numbered figure
captions with custom styling. The hook uses Pandoc markdown syntax for
images combined with OOXML for captions and paragraph formatting.

The hook delegates image rendering to Pandoc while providing enhanced
caption management with auto-numbering, cross-references, and custom
Word paragraph styles.

## Arguments

- x:

  Character string containing the path to the saved plot file.

- options:

  Named list of chunk options.

## Value

A character string containing markdown with embedded OOXML code that
combines the auto-numbered caption and the Pandoc image syntax.

## Supported chunk options

- `fig.cap`: Figure caption text. Used to generate auto-numbered
  captions with cross-reference support.

- `fig.cap.style`: Paragraph style for the caption (default from
  options).

- `fig.cap.pre`, `fig.cap.sep`: Caption prefix and separator for
  auto-numbering (e.g., "Figure" and ": ").

- `fig.id`: Cross-reference identifier (defaults to chunk label).

- `fig.topcaption`: Logical; if TRUE, caption appears above the image.

- `out.width`, `out.height`: Output dimensions. Can use units like
  "5in", "80%", "10cm". Passed directly to Pandoc.

- `out.extra`: Additional Pandoc image attributes.

- `fig.alt`: Alternative text for the image.

- `fig.align`: Image alignment. Must be one of "default", "left",
  "center", or "right". Default value "default" is converted to
  "center".

- `fig.style`: Word paragraph style name for the image paragraph
  (default: "Normal").

## Global options

- `fig.cap.tnd`: Caption numbering depth (default: 0).

- `fig.cap.tns`: Caption numbering separator (default: "-").

- `fig.cap.fp_text`: Text formatting properties for caption numbers.

## Examples of chunks

    Chunk with auto-numbered caption and custom size:

    ```{r, fig.cap="Sales over time", out.width="80%"}
    plot(sales_data)
    ```

    Chunk with alignment and custom style:

    ```{r, fig.cap="Key findings", fig.align="center", fig.style="ImageCenter"}
    plot(results)
    ```

    Caption above image with cross-reference:

    ```{r sales-plot, fig.cap="Q4 Sales", fig.topcaption=TRUE, fig.id="sales"}
    plot(sales_data)
    ```
