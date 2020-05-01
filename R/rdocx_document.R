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
#' \code{rdocx_document2} also supports cross reference based on the syntax of
#' the bookdown package.
#' @param mapstyles a named list of style to be replaced in the generated
#' document. `list("Date"="Author")` will result in a document where
#' all paragraphs styled with stylename "Date" will be styled with
#' stylename "Author".
#' @param base_format a scalar character, format to be used as a base document for
#' officedown. default to [word_document][rmarkdown::word_document] but
#' can also be word_document2 from bookdown
#' @param tables a list that can contain few items to style tables and table captions.
#' Missing items will be replaced by default values. Possible items are the following:
#'
#' * `style`: the Word stylename to use for tables. This is
#' a __table style__. You can access them in the Word template used. Function
#' [styles_info(doc, type = "table")][officer::styles_info] can let you read these
#' styles.
#' * `layout`: 'autofit' or 'fixed' algorithm. See \code{\link[officer]{table_layout}}.
#' * `width`: value of the preferred width of the table in percent (base 1).
#' * `caption`; default values `list(style = "Table Caption", pre = "Table ", sep = ": ")`
#' are producing a numbering chunk of the form "Table 2: ":
#'   * `style`: Word stylename to use for table captions.You
#' can access them in the Word template used. Function
#' [styles_info(doc, type = "paragraph")][officer::styles_info] can let you read these
#' styles.
#'   * `pre`: prefix for numbering chunk (default to "Table ").
#'   * `sep`: suffix for numbering chunk (default to ": ").
#'
#' Default value is `list(style = "Table", layout = "autofit", width = 1,
#' caption = list(style = "Table Caption", pre = "Table ", sep = ": "))`:
#'
#' ```
#' style: Table
#' layout: autofit
#' width: 1.0
#' caption:
#'   style: Table Caption
#'   pre: 'Table '
#'   sep: ': '
#' ```
#'
#' @param plots a list that can contain few items to style figures and figure captions.
#' Missing items will be replaced by default values. Possible items are the following:
#'
#' * `style`: the Word stylename to use for plots. This is a __paragraph style__.
#' You can access them in the Word template used. Function
#' [styles_info(doc, type = "paragraph")][officer::styles_info] can let you read these
#' styles.
#' * `align`: alignment of figures in the output document (possible values are 'left',
#' 'right' and 'center').
#' * `caption`; default values `list(style = "Figure Caption", pre = "Figure ", sep = ": ")`
#' are producing a numbering chunk of the form "Figure 2: ":
#'   * `style`: Word stylename to use for figure captions.You
#' can access them in the Word template used. Function
#' [styles_info(doc, type = "paragraph")][officer::styles_info] can let you read these
#' styles.
#'   * `pre`: prefix for numbering chunk (default to "Figure ").
#'   * `sep`: suffix for numbering chunk (default to ": ").
#'
#' Default value is `list(style = "Normal", align = "center",
#' caption = list(style = "Image Caption", pre = "Figure ", sep = ": "))`:
#'
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
#' @param ... arguments used by [word_document][rmarkdown::word_document]
#'
#'
#'
#'
#' @section table style:
#'
#' Pandoc does not allow usage of Word table style. This option
#' allows you to define which Word table style is the default.
#' These table styles must be present in the `reference_docx` document.
#' It can be read with [styles_info(doc, type = "table")][officer::styles_info]
#' or within Word table styles view.
#'
#' To create a table style in your `reference_docx` corresponding to your needs,
#' edit the document with MS Word and add a new style of type "table" then configure
#' it. The style name must be used as the value of the "tab.style" argument.
#'
#' \if{html}{
#'
#' You should see a window that looks like the one below:
#'
#' \figure{TABLES-new-style.png}{options: width=400px}
#'
#' In the Define New Table Style window, start give your new style a name.
#' There are a many formatting options available in this window.
#' For example, you can change the font and font style, change the border
#' and cell colors, and change the text alignment.
#'
#' }
#'
#' The package is only using these styles and is not able to create them with
#' R code.
#' @section lists:
#' Pandoc does not allow easy customization of ordered or unordered lists. This option
#' allows you to apply a list style for ordered lists and a list style for unordered
#' lists. These list styles must be present in the `reference_docx` document.
#'
#' To create a list style in your `reference_docx` corresponding to your needs,
#' edit the document with MS Word and add a new style of type "list" then configure
#' it. The style name must be used as the value of the "ol.style" argument if you
#' configure an ordered list (i.e. with numbers corresponding to each level) or
#' as the value of the "ul.style" argument if you configure an unordered list
#' (i.e. with bullets corresponding to each level).
#'
#' \if{html}{
#'
#' You should see a window that looks like the one below:
#'
#' \figure{LISTS-new_style.png}{options: width=400px}
#'
#' In the Define New List Style window, start give your new style a name.
#' There are a many formatting options available in this window. You can
#' change the font, define the character formatting and choose the
#' type (number or bullet).
#'
#' }
#'
#' The package is only using these styles and is not able to create them with
#' R code.
#' @examples
#' library(rmarkdown)
#'
#' # official template -----
#' skeleton <- system.file(package = "officedown",
#'   "rmarkdown/templates/word/skeleton/skeleton.Rmd")
#' rmd_file <- tempfile(fileext = ".Rmd")
#' file.copy(skeleton, to = rmd_file)
#'
#' docx_file_1 <- tempfile(fileext = ".docx")
#' render(rmd_file, output_file = docx_file_1, quiet = TRUE)
#'
#' # bookdown example -----
#'
#' # All above is only to make sure we do not write in your wd
#' bookdown_loc <- system.file(package = "officedown", "examples/bookdown")
#' temp_dir <- tempfile()
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
#' @importFrom officer change_styles
#' @importFrom utils modifyList
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
#' ---
#' ```
rdocx_document <- function(mapstyles,
                           base_format = "rmarkdown::word_document",
                           tables = list(), plots = list(), lists = list(),
                           ...) {

  base_format_fun <- get_fun(base_format)
  output_formats <- base_format_fun(...)


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

  if( missing(mapstyles) )
    mapstyles <- list()

  output_formats$post_knit <- function(metadata, input_file, runtime, ...){
    output_file <- file_with_meta_ext(input_file, "knit", "md")
    content <- readLines(output_file)

    content <- post_knit_table_captions( content,
      tab.cap.pre = tables$caption$pre, tab.cap.sep = tables$caption$sep)
    content <- post_knit_references(content, lp = "tab:")
    content <- post_knit_references(content, lp = "fig:")
    content <- post_knit_references(content)
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
    x <- change_styles(x, mapstyles = mapstyles)
    forget(get_reference_rdocx)
    print(x, target = output_file)
    output_file
  }
  output_formats$bookdown_output_format = 'docx'
  output_formats

}

#' @rdname rdocx_document
#' @importFrom bookdown markdown_document2
#' @export
rdocx_document2 <- function(...) {
  markdown_document2(..., base_format = rdocx_document)
}



