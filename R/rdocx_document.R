# utils ----
#' @importFrom utils getAnywhere getFromNamespace
get_fun <- function(x){
  if( grepl("::", x, fixed = TRUE) ){
    coumpounds <- strsplit(x, split = "::", x, fixed = TRUE)[[1]]
    z <- getFromNamespace(coumpounds[2], ns = coumpounds[1] )
  } else {
    z <- getAnywhere(x)
    if(length(z$objs) < 1){
      stop("could not find any function named ", shQuote(z$name), " in loaded namespaces or in the search path. If the package is installed, specify name with `packagename::function_name`.")
    }
  }
  z
}

file_with_meta_ext <- function(file, meta_ext, ext = tools::file_ext(file)) {
  paste(tools::file_path_sans_ext(file),
        ".", meta_ext, ".", ext,
        sep = ""
  )
}

# tables_default_values ----
tables_default_values <- list(
  style = "Table",
  layout = "autofit",
  width = 1,
  caption = list(
    style = "Table Caption",
    pre = "Table ", sep = ": "
  ),
  conditional = list(
    first_row = TRUE,
    first_column = FALSE,
    last_row = FALSE,
    last_column = FALSE,
    no_hband = FALSE,
    no_vband = TRUE
    )
)

# plots_default_values ----
plots_default_values <- list(
  style = "Figure",
  align = "center",
  caption = list(
    style = "Image Caption",
    pre = "Figure ", sep = ": "
  )
)

# lists_default_values ----
lists_default_values <- list(
  ol.style = NULL,
  ul.style = NULL
)

# memoise reference_docx ----

#' @importFrom officer get_reference_value
get_docx_uncached <- function() {
  ref_docx <- read_docx(get_reference_value(format = "docx"))
  ref_docx
}

#' @importFrom memoise memoise forget
get_reference_rdocx <- memoise(get_docx_uncached)


# main ----
#' @export
#' @title Advanced R Markdown Word Format
#' @description Format for converting from R Markdown to an MS Word
#' document. The function comes also with improved output options.
#' @param base_format a scalar character, format to be used as a base document for
#' officedown. default to [word_document][rmarkdown::word_document] but
#' can also be word_document2 from bookdown
#' @param tables a list that can contain few items to style tables and table captions.
#' Missing items will be replaced by default values. Possible items are the following:
#'
#' * `style`: the Word stylename to use for tables.
#' * `layout`: 'autofit' or 'fixed' algorithm. See \code{\link[officer]{table_layout}}.
#' * `width`: value of the preferred width of the table in percent (base 1).
#' * `caption`; caption options, i.e.:
#'   * `style`: Word stylename to use for table captions.
#'   * `pre`: prefix for numbering chunk (default to "Table ").
#'   * `sep`: suffix for numbering chunk (default to ": ").
#' * `conditional`: a list of named logical values:
#'   * `first_row` and `last_row`: apply or remove formatting from the first or last row in the table
#'   * `first_column`  and `last_column`: apply or remove formatting from the first or last column in the table
#'   * `no_hband` and `no_vband`: don't display odd and even rows or columns with alternating shading for ease of reading.
#'
#' Default value is (in R format):
#' ```
#' list(
#'    style = "Table", layout = "autofit", width = 1,
#'    caption = list(
#'      style = "Table Caption", pre = "Table ", sep = ": "),
#'    conditional = list(
#'      first_row = TRUE, first_column = FALSE, last_row = FALSE,
#'      last_column = FALSE, no_hband = FALSE, no_vband = TRUE
#'    )
#' )
#' ```
#'
#' Default value is (in YAML format):
#' ```
#' style: Table
#' layout: autofit
#' width: 1.0
#' caption:
#'   style: Table Caption
#'   pre: 'Table '
#'   sep: ': '
#' conditional:
#'   first_row: true
#'   first_column: false
#'   last_row: false
#'   last_column: false
#'   no_hband: false
#'   no_vband: true
#' ```
#'
#' @param plots a list that can contain few items to style figures and figure captions.
#' Missing items will be replaced by default values. Possible items are the following:
#'
#' * `style`: the Word stylename to use for plots.
#' * `align`: alignment of figures in the output document (possible values are 'left',
#' 'right' and 'center').
#' * `caption`; caption options, i.e.:
#'   * `style`: Word stylename to use for figure captions.
#'   * `pre`: prefix for numbering chunk (default to "Figure ").
#'   * `sep`: suffix for numbering chunk (default to ": ").
#'
#' Default value is (in R format):
#' ```
#' list(
#'   style = "Normal", align = "center",
#'   caption = list(
#'     style = "Image Caption",
#'     pre = "Figure ",
#'     sep = ": "
#'    )
#'  )
#'  ```
#'
#' Default value is (in YAML format):
#' ```
#' style: Normal
#' align: center
#' caption:
#'   style: Image Caption
#'   pre: 'Figure '
#'   sep: ': '
#' ```
#' @param lists a list containing two named items `ol.style` and
#' `ul.style`, values are the stylenames to be used to replace the style of ordered
#' and unordered lists created by pandoc. If NULL, no replacement is made.
#'
#' Default value is `list(ol.style = NULL, ul.style = NULL)`:
#'
#' ```
#' ol.style: null
#' ul.style: null
#' ```
#' @param mapstyles a named list of style to be replaced in the generated
#' document. `list("Normal" = c("Author", "Date"))` will result in a document where
#' all paragraphs styled with stylename "Date" and "Author" will be then styled with
#' stylename "Normal".
#' @param reference_num if TRUE, text for references to sections will be
#' the section number (e.g. '3.2'). If FALSE, text for references to sections
#' will be the text (e.g. 'section title').
#' @param ... arguments used by [word_document][rmarkdown::word_document]
#' @return R Markdown output format to pass to [render][rmarkdown::render]
#' @section Finding stylenames:
#'
#' You can access them in the Word template used. Function
#' [styles_info()][officer::styles_info] can let you read these
#' styles.
#'
#' You need officer to read the stylenames (to get information
#' from a specific "reference_docx", change `ref_docx_default`
#' in the example below.
#'
#' ```
#' library(officer)
#' docx_file <- system.file(package = "officer", "template", "template.docx")
#' doc <- read_docx(docx_file)
#' ```
#'
#' To read `paragraph` stylenames:
#' ```
#' styles_info(doc, type = "paragraph")
#' ```
#'
#' To read `table` stylenames:
#' ```
#' styles_info(doc, type = "table")
#' ```
#'
#' To read `list` stylenames:
#' ```
#' styles_info(doc, type = "numbering")
#' ```
#'
#'
#' @section R Markdown yaml:
#' The following demonstrates how to pass arguments in the R Markdown yaml:
#'
#' ```
#' ---
#' output:
#'   officedown::rdocx_document:
#'     reference_docx: pandoc_template.docx
#'     tables:
#'       style: Table
#'       layout: autofit
#'       width: 1.0
#'       caption:
#'         style: Table Caption
#'         pre: 'Table '
#'         sep: ': '
#'       conditional:
#'         first_row: true
#'         first_column: false
#'         last_row: false
#'         last_column: false
#'         no_hband: false
#'         no_vband: true
#'     plots:
#'       style: Normal
#'       align: center
#'       caption:
#'         style: Image Caption
#'         pre: 'Figure '
#'         sep: ': '
#'     lists:
#'       ol.style: null
#'       ul.style: null
#'     mapstyles:
#'       Normal: ['First Paragraph', 'Author', 'Date']
#'     reference_num: true
#' ---
#' ```
#' @examples
#' library(rmarkdown)
#' run_ok <- pandoc_available() &&
#'   pandoc_version() >= numeric_version("2.0")
#'
#' if(run_ok){
#'
#' # minimal example -----
#' example <- system.file(package = "officedown",
#'   "examples/minimal_word.Rmd")
#' rmd_file <- tempfile(fileext = ".Rmd")
#' file.copy(example, to = rmd_file)
#'
#' docx_file_1 <- tempfile(fileext = ".docx")
#' render(rmd_file, output_file = docx_file_1, quiet = TRUE)
#'
#' # bookdown example -----
#'
#' bookdown_loc <- system.file(package = "officedown", "examples/bookdown")
#'
#' temp_dir <- tempfile()
#' # uncomment next line to get the result in your working directory
#' # temp_dir <- "./bd_example"
#'
#' dir.create(temp_dir, showWarnings = FALSE, recursive = TRUE)
#' file.copy(
#'   from = list.files(bookdown_loc, full.names = TRUE),
#'   to = temp_dir,
#'   overwrite = TRUE, recursive = TRUE)
#'
#' render_site(
#'   input = temp_dir, encoding = 'UTF-8',
#'   envir = new.env(), quiet = TRUE)
#'
#' docx_file_2 <- file.path(temp_dir, "_book", "bookdown.docx")
#'
#' if(file.exists(docx_file_2)){
#'   message("file ", docx_file_2, " has been written.")
#' }
#' }
#' @importFrom officer change_styles
#' @importFrom utils modifyList
rdocx_document <- function(base_format = "rmarkdown::word_document",
                           tables = list(), plots = list(), lists = list(),
                           mapstyles = list(),
                           reference_num = TRUE,
                           ...) {

  args <- list(...)
  if(is.null(args$reference_docx)){
    args$reference_docx <- system.file(
      package = "officedown", "examples",
      "bookdown", "template.docx"
    )
  }

  base_format_fun <- get_fun(base_format)
  output_formats <- do.call(base_format_fun, args)


  tables <- modifyList(tables_default_values, tables)
  plots <- modifyList(plots_default_values, plots)
  lists <- modifyList(lists_default_values, lists)

  output_formats$knitr$opts_chunk <- append(
    output_formats$knitr$opts_chunk,
    list(tab.cap.style = tables$caption$style,
         tab.cap.pre = tables$caption$pre,
         tab.cap.sep = tables$caption$sep,
         tab.lp = "tab:",
         tab.style = tables$style,
         tab.width = tables$width,

         first_row = tables$conditional$first_row,
         first_column = tables$conditional$first_column,
         last_row = tables$conditional$last_row,
         last_column = tables$conditional$last_column,
         no_hband = tables$conditional$no_hband,
         no_vband = tables$conditional$no_vband,

         fig.cap.style = plots$caption$style,
         fig.cap.pre = plots$caption$pre,
         fig.cap.sep = plots$caption$sep,
         fig.align = plots$align,
         fig.style = plots$style,
         fig.lp = "fig:"
         )
    )
  if(is.null(output_formats$knitr$knit_hooks)){
    output_formats$knitr$knit_hooks <- list()
  }
  output_formats$knitr$knit_hooks$plot <- plot_word_fig_caption

  output_formats$post_knit <- function(metadata, input_file, runtime, ...){
    output_file <- file_with_meta_ext(input_file, "knit", "md")
    content <- readLines(output_file)

    content <- post_knit_table_captions(content,
      tab.cap.pre = tables$caption$pre, tab.cap.sep = tables$caption$sep,
      style = tables$caption$style)
    content <- post_knit_caption_references(content, lp = "tab:")
    content <- post_knit_caption_references(content, lp = "fig:")
    content <- post_knit_std_references(content, numbered = reference_num)
    content <- block_macro(content)
    writeLines(content, output_file)
  }

  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_docx(output_file)
    x <- process_images(x)
    x <- process_links(x)
    x <- process_embedded_docx(x)
    x <- process_par_settings(x)
    x <- process_list_settings(x, ul_style = lists$ul.style, ol_style = lists$ol.style)
    x <- process_sections(x)
    x <- change_styles(x, mapstyles = mapstyles)
    forget(get_reference_rdocx)
    print(x, target = output_file)
    output_file
  }
  output_formats$bookdown_output_format = 'docx'
  output_formats

}
