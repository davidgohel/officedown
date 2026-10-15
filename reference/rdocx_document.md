# Advanced R Markdown Word Format

'R Markdown' Format for converting from 'R Markdown' document to an MS
Word document.

The function enhances the output offered by
[`rmarkdown::word_document()`](https://pkgs.rstudio.com/rmarkdown/reference/word_document.html)
with advanced formatting features.

## Usage

``` r
rdocx_document(
  base_format = "rmarkdown::word_document",
  tables = list(),
  plots = list(),
  lists = list(),
  mapstyles = list(),
  page_size = NULL,
  page_margins = NULL,
  reference_num = TRUE,
  ...
)
```

## Arguments

- base_format:

  a scalar character, the format to be used as a base document for
  'officedown'. Default to
  [word_document](https://pkgs.rstudio.com/rmarkdown/reference/word_document.html)
  but can also be `word_document2()` from bookdown.

  When the `base_format` used is
  [`bookdown::word_document2`](https://pkgs.rstudio.com/bookdown/reference/html_document2.html),
  the `number_sections` parameter is automatically set to `FALSE`.
  Indeed, if you want numbered titles, you are asked to use a Word
  document template with auto-numbered titles (the title styles of the
  default \`rdocx_document' template are already set to FALSE).

- tables:

  see section 'Tables' below.

- plots:

  see section 'Plots' below.

- lists:

  see section 'Lists' below.

- mapstyles:

  a named list of style to be replaced in the generated document.
  `list("Normal" = c("Author", "Date"))` will result in a document where
  all paragraphs styled with stylename "Date" and "Author" will be then
  styled with stylename "Normal".

- page_size, page_margins:

  default page and margins dimensions. If not null (the default), these
  values are used to define the default Word section. See
  [`officer::page_size()`](https://davidgohel.github.io/officer/reference/page_size.html)
  and
  [`officer::page_mar()`](https://davidgohel.github.io/officer/reference/page_mar.html).

- reference_num:

  if `TRUE`, text for references to sections will be the section number
  (e.g. '3.2'). If FALSE, text for references to sections will be the
  text (e.g. 'section title').

- ...:

  arguments used by
  [word_document](https://pkgs.rstudio.com/rmarkdown/reference/word_document.html)

## Value

R Markdown *output format* to pass to
[render](https://pkgs.rstudio.com/rmarkdown/reference/render.html).

## Tables

a list that can contain few items to style tables and table captions.
Missing items will be replaced by default values. Possible items are the
following:

- `style`: the Word stylename to use for tables.

- `layout`: 'autofit' or 'fixed' algorithm. See
  [`table_layout`](https://davidgohel.github.io/officer/reference/table_layout.html).

- `width`: value of the preferred width of the table in percent (base
  1).

- `topcaption`: caption will appear before (on top of) the table,

- `tab.lp`: caption table sequence identifier. All table captions are
  supposed to have the same identifier. It makes possible to insert list
  of tables. It is also used to prefix your 'bookdown' cross-reference
  call; if `tab.lp` is set to "tab:", a cross-reference to table with id
  "xxxxx" is written as `\@ref(tab:xxxxx)`. It is possible to set the
  value to your default Word value (in French for example it is
  "Tableau", in German it is "Tabelle"), you can then add manually a
  list of tables (go to the "References" tab and select menu "Insert
  Table of Figures").

- `caption`; caption options, i.e.:

  - `style`: Word stylename to use for table captions.

  - `pre`: prefix for numbering chunk (default to "Table ").

  - `sep`: suffix for numbering chunk (default to ": ").

  - `tnd`: (only applies if positive. )Inserts the number of the last
    title of level `tnd` (i.e. 4.3-2 for figure 2 of chapter 4.3).

  - `tns`: separator to use between title number and table number.
    Default is "-".

  - `fp_text`: text formatting properties to apply to caption prefix -
    see
    [`officer::fp_text_lite()`](https://davidgohel.github.io/officer/reference/fp_text.html).

- `conditional`: a list of named logical values:

  - `first_row` and `last_row`: apply or remove formatting from the
    first or last row in the table

  - `first_column` and `last_column`: apply or remove formatting from
    the first or last column in the table

  - `no_hband` and `no_vband`: don't display odd and even rows or
    columns with alternating shading for ease of reading.

To cross-reference a table, give it an identifier with the chunk option
`tab.id`, along with a caption (chunk option `tab.cap`); unless
'bookdown' is used, the chunk label is not used for tables. A table with
`tab.id = "mtcars"` is then referenced with `\@ref(tab:mtcars)`. Without
`tab.id`, no bookmark is created and Word displays "Error! Reference
source not found." in place of the reference. Identifiers can contain
letters, digits, `-` and `_`.

Default value is (in YAML format):

    style: Table
    layout: autofit
    width: 1.0
    topcaption: true
    tab.lp: 'tab:'
    caption:
      style: Table Caption
      pre: 'Table'
      sep: ':'
      tnd: 0
      tns: '-'
      fp_text: !expr officer::fp_text_lite(bold = TRUE)
    conditional:
      first_row: true
      first_column: false
      last_row: false
      last_column: false
      no_hband: false
      no_vband: true

## Plots

Argument plot is expected to be a list. It can contain few items to
style figures and figure captions.

You don't have to provide values for each items of the list, missing
items are replaced by default values. Possible items are:

- `style`: the name of the "Word" paragraph style to use for the
  paragraphs that will contain a graphic.

- `align`: alignment of paragraphs containing graphics (possible values
  are 'left', 'right' and 'center').

- `topcaption`: specify whether the legend should appear above the
  graphic (`TRUE`) or below the graphic (`FALSE`).

- `fig.lp`: Sequence identifier for figure legends. All graphics'
  captions are supposed to have the same identifier. This value is used
  for different features:

  - It is used to insert a table of figures.

  - It is also used to prefix your 'bookdown' cross-reference call; if
    `fig.lp` is set to "fig:", a cross-reference to the figure with id
    "xxxxx" is written as `\@ref(fig:xxxxx)`.

  It is possible to set the Word default value (in French for example,
  it is "Figure"), you can then manually add a list of figures (go to
  the "References" tab and select the "Insert a table of figures" menu).

- `caption`: a list for captions' options, i.e.

  - `style`: the name of the "Word" paragraph style to use for the
    paragraphs that will contain figures' captions.

  - `pre`: prefix for numbering chunk (default to "Figure ").

  - `sep`: suffix for numbering chunk (default to ": ").

  - `tnd`: (only applies if positive). Inserts the number of the last
    title of level `tnd` (i.e. 4.3-2 for figure 2 of chapter 4.3).

  - `tns`: separator to use between title number and figure number.
    Default is "-".

  - `fp_text`: text formatting properties to apply to caption prefix -
    see
    [`officer::fp_text_lite()`](https://davidgohel.github.io/officer/reference/fp_text.html).

To cross-reference a figure, give it a caption (chunk option `fig.cap`).
Its identifier is the chunk label, unless the chunk option `fig.id` is
set. A figure produced by a chunk labelled `boxplot` is then referenced
with `\@ref(fig:boxplot)`. Identifiers can contain letters, digits, `-`
and `_`; a chunk label containing `.` and a caption makes the rendering
fail, use `fig.id` to provide a valid identifier.

Default value is (in YAML format):

    style: Normal
    align: center
    topcaption: false
    fig.lp: 'fig:'
    caption:
      style: Image Caption
      pre: 'Figure '
      sep: ': '
      tnd: 0
      tns: '-'
      fp_text: !expr officer::fp_text_lite(bold = TRUE)

## Lists

The parameter lists is a list that can contain two named items:

- `ol.style`: style to use for ordered lists.

- `ul.style`: style to use for unordered lists.

Default values are `list(ol.style = NULL, ul.style = NULL)`.

Expected values are the stylenames to be used to replace the style of
ordered and unordered lists created by pandoc. If NULL, no replacement
is made.

These values in YAML format are:

    output:
      officedown::rdocx_document:
        lists:
          ol.style: null
          ul.style: null

They can have values corresponding to existing stylenames (of type
'numbering'). With package 'officer', we can read these values with
[styles_info()](https://davidgohel.github.io/officer/reference/styles_info.html).

    library(officer)
    docx_file <- system.file(
      package = "officedown", "examples",
      "bookdown", "template.docx"
    )
    doc <- read_docx(docx_file)
    styles_info(doc, type = "numbering")[, 1:6]
    #>    style_type    style_id style_name     base_on is_custom is_default
    #> 13  numbering Aucuneliste    No List        <NA>     FALSE       TRUE
    #> 40  numbering   Defaultul Default ul Aucuneliste      TRUE      FALSE
    #> 41  numbering   Defaultol Default ol Aucuneliste      TRUE      FALSE

From the above available values, the possible configuration is possible:

    output:
      officedown::rdocx_document:
        lists:
          ol.style: 'Default ol'
          ul.style: 'Default ul'

## Finding stylenames

You can access them in the Word template used. Function
[styles_info()](https://davidgohel.github.io/officer/reference/styles_info.html)
can let you read these styles.

You need 'officer' to read the stylenames (to get information from a
specific "reference_docx", change `ref_docx_default` in the example
below.

    library(officer)
    docx_file <- system.file(package = "officer", "template", "template.docx")
    doc <- read_docx(docx_file)

To read `paragraph` stylenames:

    styles_info(doc, type = "paragraph")[, 1:6]
    #>    style_type      style_id    style_name      base_on is_custom is_default
    #> 1   paragraph        Normal        Normal         <NA>     FALSE       TRUE
    #> 2   paragraph        Titre1     heading 1       Normal     FALSE      FALSE
    #> 3   paragraph        Titre2     heading 2       Normal     FALSE      FALSE
    #> 4   paragraph        Titre3     heading 3       Normal     FALSE      FALSE
    #> 9   paragraph      centered      centered       Normal      TRUE      FALSE
    #> 15  paragraph  ImageCaption Image Caption       Normal      TRUE      FALSE
    #> 16  paragraph  TableCaption Table Caption ImageCaption      TRUE      FALSE
    #> 18  paragraph           TM1         toc 1       Normal     FALSE      FALSE
    #> 19  paragraph           TM2         toc 2       Normal     FALSE      FALSE
    #> 20  paragraph Textedebulles  Balloon Text       Normal     FALSE      FALSE
    #> 23  paragraph  graphictitle graphic title ImageCaption      TRUE      FALSE
    #> 24  paragraph    tabletitle   table title TableCaption      TRUE      FALSE

To read `table` stylenames:

    styles_info(doc, type = "table")[, 1:6]
    #>    style_type             style_id          style_name       base_on is_custom
    #> 6       table        TableauNormal        Normal Table          <NA>     FALSE
    #> 10      table        tabletemplate      table_template TableauNormal      TRUE
    #> 11      table  Listeclaire-Accent2 Light List Accent 2 TableauNormal     FALSE
    #> 17      table Tableauprofessionnel  Table Professional TableauNormal     FALSE
    #>    is_default
    #> 6        TRUE
    #> 10      FALSE
    #> 11      FALSE
    #> 17      FALSE

To read `list` stylenames:

    styles_info(doc, type = "numbering")[, 1:6]
    #>   style_type    style_id style_name base_on is_custom is_default
    #> 7  numbering Aucuneliste    No List    <NA>     FALSE       TRUE

## R Markdown yaml

The following demonstrates how to pass arguments in the R Markdown yaml:

    ---
    output:
      officedown::rdocx_document:
        reference_docx: pandoc_template.docx
        tables:
          style: Table
          layout: autofit
          width: 1.0
          topcaption: true
          tab.lp: 'tab:'
          caption:
            style: Table Caption
            pre: 'Table '
            sep: ': '
            tnd: 0
            tns: '-'
            fp_text: !expr officer::fp_text_lite(bold = TRUE)
          conditional:
            first_row: true
            first_column: false
            last_row: false
            last_column: false
            no_hband: false
            no_vband: true
        plots:
          style: Normal
          align: center
          fig.lp: 'fig:'
          topcaption: false
          caption:
            style: Image Caption
            pre: 'Figure '
            sep: ': '
            tnd: 0
            tns: '-'
            fp_text: !expr officer::fp_text_lite(bold = TRUE)
        lists:
          ol.style: null
          ul.style: null
        mapstyles:
          Normal: ['First Paragraph', 'Author', 'Date']
        page_size:
          width: 11906 / 1440
          height: 16838 / 1440
          orient: "portrait"
        page_margins:
          bottom: 1
          top: 1
          right: 1417 / 1440
          left: 1417 / 1440
          header: 708 / 1440
          footer: 708 / 1440
          gutter: 0 / 1440
        reference_num: true
    ---

## Known limitations

When using
[`knitr::include_graphics()`](https://rdrr.io/pkg/knitr/man/include_graphics.html)
in 'rdocx_document', the image must be local, while with standard 'R
Markdown' output documents, images located on Internet are supported.

## Examples

``` r
# rdocx_document basic example -----
library(rmarkdown)
library(officedown)

if (pandoc_available() && pandoc_version() >= numeric_version("2.0")) {
  # minimal example -----
  example <- system.file(
    package = "officedown",
    "examples/minimal_word.Rmd"
  )
  rmd_file <- tempfile(fileext = ".Rmd")
  file.copy(example, to = rmd_file)

  docx_file <- tempfile(fileext = ".docx")
  render(rmd_file, output_file = docx_file, quiet = TRUE)
}
```
