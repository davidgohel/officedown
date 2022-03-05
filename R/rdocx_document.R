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

absolute_path <- function(x){
  if (length(x) != 1L)
    stop("'x' must be a single character string")
  epath <- path.expand(x)
  if( file.exists(epath)){
    epath <- normalizePath(epath, "/", mustWork = TRUE)
  } else {
    if( !dir.exists(dirname(epath)) ){
      stop("directory of ", x, " does not exist.", call. = FALSE)
    }
    cat("", file = epath)
    epath <- normalizePath(epath, "/", mustWork = TRUE)
    unlink(epath)
  }
  epath
}

# tables_default_values ----
tables_default_values <- list(
  style = "Table",
  layout = "autofit",
  width = 1,
  tab.lp = "tab:",
  topcaption = TRUE,
  caption = list(
    style = "Table Caption",
    pre = "Table ", sep = ": ",
    tnd = 0,
    tns = "-",
    fp_text = fp_text_lite(bold = TRUE)
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
  fig.lp = "fig:",
  topcaption = FALSE,
  caption = list(
    style = "Image Caption",
    pre = "Figure ", sep = ": ",
    tnd = 0,
    tns = "-",
    fp_text = fp_text_lite(bold = TRUE)
  )
)

# lists_default_values ----
lists_default_values <- list(
  ol.style = NULL,
  ul.style = NULL
)

# page_size_default_values ----
page_size_default_values <- list(
  width = 8.3,
  height = 11.7,
  orient = "portrait"
)
# page_mar_default_values ----
page_mar_default_values <- list(
  bottom = 1.25,
  top = 1.25,
  right = 0.5,
  left = .5,
  header = 0.5,
  footer = 0.5,
  gutter = 0.5
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
#' can also be word_document2 from bookdown.
#'
#' When the `base_format` used is `bookdown::word_document2`, the `number_sections`
#' parameter is automatically set to `FALSE`. Indeed, if you want numbered titles,
#' you are asked to use a Word document template with auto-numbered titles (the title
#' styles of the default `rdocx_document' template are already set to FALSE).
#'
#' @param tables a list that can contain few items to style tables and table captions.
#' Missing items will be replaced by default values. Possible items are the following:
#'
#' * `style`: the Word stylename to use for tables.
#' * `layout`: 'autofit' or 'fixed' algorithm. See \code{\link[officer]{table_layout}}.
#' * `width`: value of the preferred width of the table in percent (base 1).
#' * `topcaption`: caption will appear before (on top of) the table,
#' * `tab.lp`: caption table sequence identifier. All table captions are supposed
#' to have the same identifier. It makes possible to insert list of tables. It is
#' also used to prefix your 'bookdown' cross-reference call; if `tab.lp` is set to
#' "tab:", a cross-reference to table with id "xxxxx" is written as `\@ref(tab:xxxxx)`.
#' It is possible to set the value to your default Word value (in French for example it
#' is "Tableau", in German it is "Tabelle"), you can then add manually a list of
#' tables (go to the "References" tab and select menu "Insert Table of Figures").
#' * `caption`; caption options, i.e.:
#'   * `style`: Word stylename to use for table captions.
#'   * `pre`: prefix for numbering chunk (default to "Table ").
#'   * `sep`: suffix for numbering chunk (default to ": ").
#'   * `tnd`: (only applies if positive. )Inserts the number of the last title of level `tnd` (i.e. 4.3-2 for figure 2 of chapter 4.3).
#'   * `tns`: separator to use between title number and table number. Default is "-".
#'   * `fp_text`: text formatting properties to apply to caption prefix - see [fp_text_lite()].
#' * `conditional`: a list of named logical values:
#'   * `first_row` and `last_row`: apply or remove formatting from the first or last row in the table
#'   * `first_column`  and `last_column`: apply or remove formatting from the first or last column in the table
#'   * `no_hband` and `no_vband`: don't display odd and even rows or columns with alternating shading for ease of reading.
#'
#'
#' Default value is (in YAML format):
#' ```
#' style: Table
#' layout: autofit
#' width: 1.0
#' topcaption: true
#' tab.lp: 'tab:'
#' caption:
#'   style: Table Caption
#'   pre: 'Table '
#'   sep: ': '
#'   tnd: 0
#'   tns: '-'
#'   fp_text: !expr officer::fp_text_lite(bold = TRUE)
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
#' * `topcaption`: caption will appear before (on top of) the figure,
#' * `fig.lp`: caption figure sequence identifier. All figure captions are supposed
#' to have the same identifier. It makes possible to insert list of figures. It is
#' also used to prefix your 'bookdown' cross-reference call; if `fig.lp` is set to
#' "fig:", a cross-reference to figure with id "xxxxx" is written as `\@ref(fig:xxxxx)`.
#' It is possible to set the value to your default Word value (in French for example it
#' is "Figure"), you can then add manually a list of
#' figures (go to the "References" tab and select menu "Insert Table of Figures").
#' * `caption`; caption options, i.e.:
#'   * `style`: Word stylename to use for figure captions.
#'   * `pre`: prefix for numbering chunk (default to "Figure ").
#'   * `sep`: suffix for numbering chunk (default to ": ").
#'   * `tnd`: (only applies if positive. )Inserts the number of the last title of level `tnd` (i.e. 4.3-2 for figure 2 of chapter 4.3).
#'   * `tns`: separator to use between title number and figure number. Default is "-".
#'   * `fp_text`: text formatting properties to apply to caption prefix - see [fp_text_lite()].
#'
#'
#' Default value is (in YAML format):
#' ```
#' style: Normal
#' align: center
#' topcaption: false
#' fig.lp: 'fig:'
#' caption:
#'   style: Image Caption
#'   pre: 'Figure '
#'   sep: ': '
#'   tnd: 0
#'   tns: '-'
#'   fp_text: !expr officer::fp_text_lite(bold = TRUE)
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
#' @param page_size,page_margins default page and margins dimensions, these values are used to define the default Word section.
#' See [page_size()] and [page_mar()].
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
#'       topcaption: true
#'       tab.lp: 'tab:'
#'       caption:
#'         style: Table Caption
#'         pre: 'Table '
#'         sep: ': '
#'         tnd: 0
#'         tns: '-'
#'         fp_text: !expr officer::fp_text_lite(bold = TRUE)
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
#'       fig.lp: 'fig:'
#'       topcaption: false
#'       caption:
#'         style: Image Caption
#'         pre: 'Figure '
#'         sep: ': '
#'         tnd: 0
#'         tns: '-'
#'         fp_text: !expr officer::fp_text_lite(bold = TRUE)
#'     lists:
#'       ol.style: null
#'       ul.style: null
#'     mapstyles:
#'       Normal: ['First Paragraph', 'Author', 'Date']
#'     page_size:
#'       width: 8.3
#'       height: 11.7
#'       orient: "portrait"
#'     page_margins:
#'       bottom: 1
#'       top: 1
#'       right: 1.25
#'       left: 1.25
#'       header: 0.5
#'       footer: 0.5
#'       gutter: 0.5
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
#' render(rmd_file, output_file = docx_file_1, quiet = TRUE,
#'   intermediates_dir = tempfile())
#'
#' # bookdown example -----
#' if(require("bookdown")){
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
#'
#' }
#' @importFrom officer change_styles
#' @importFrom utils modifyList
rdocx_document <- function(base_format = "rmarkdown::word_document",
                           tables = list(), plots = list(), lists = list(),
                           mapstyles = list(), page_size = list(), page_margins = list(),
                           reference_num = TRUE, ...) {

  args <- list(...)
  if(is.null(args$reference_docx)){
    args$reference_docx <- system.file(
      package = "officedown", "examples",
      "bookdown", "template.docx"
    )
  }
  if(!is.null(args$number_sections) && isTRUE(args$number_sections)){
    args$number_sections <- FALSE
  }

  args$reference_docx <- absolute_path(args$reference_docx)

  base_format_fun <- get_fun(base_format)
  output_formats <- do.call(base_format_fun, args)


  tables <- modifyList(tables_default_values, tables)
  plots <- modifyList(plots_default_values, plots)
  lists <- modifyList(lists_default_values, lists)
  page_size <- modifyList(page_size_default_values, page_size)
  page_margins <- modifyList(page_mar_default_values, page_margins)

  output_formats$knitr$opts_chunk <- append(
    output_formats$knitr$opts_chunk,
    list(tab.cap.style = tables$caption$style,
         tab.cap.pre = tables$caption$pre,
         tab.cap.sep = tables$caption$sep,
         tab.cap.tnd = tables$caption$tnd,
         tab.cap.tns = tables$caption$tns,
         tab.cap.fp_text = tables$caption$fp_text,
         tab.lp = tables$tab.lp,
         tab.topcaption = tables$topcaption,
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
         fig.cap.tnd = plots$caption$tnd,
         fig.cap.tns = plots$caption$tns,
         fig.cap.fp_text = plots$caption$fp_text,
         fig.align = plots$align,
         fig.style = plots$style,
         fig.lp = plots$fig.lp,
         fig.topcaption = plots$topcaption
         )
    )
  if(is.null(output_formats$knitr$knit_hooks)){
    output_formats$knitr$knit_hooks <- list()
  }
  output_formats$knitr$knit_hooks$plot <- plot_word_fig_caption

  # This is the intermediate_dir and wll be updated if `intermediates_generator` is called
  intermediate_dir <- "."

  temp_intermediates_generator <- output_formats$intermediates_generator
  output_formats$intermediates_generator <- function(...){
    intermediate_dir <<- list(...)[[2]]
    temp_intermediates_generator(...)
  }

  output_formats$post_knit <- function(
    metadata, input_file, runtime, ...){
    output_file <- file_with_meta_ext(input_file, "knit", "md")
    output_file <- file.path(intermediate_dir, output_file)
    content <- readLines(output_file)

    # content <- post_knit_table_captions(
    #   content = content,
    #   tab.cap.pre = tables$caption$pre,
    #   tab.cap.sep = tables$caption$sep,
    #   style = tables$caption$style,
    #   tab.lp = tables$tab.lp,
    #   tnd = tables$caption$tnd,
    #   tns = tables$caption$tns,
    #   prop = tables$caption$fp_text
    # )

    content <- post_knit_caption_references(content, lp = tables$tab.lp)
    content <- post_knit_caption_references(content, lp = plots$fig.lp)
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
    x <- change_styles(x, mapstyles = mapstyles)

    # default section
    default_sect_properties <- prop_section(
      page_size = page_size(
        orient = page_size$orient,
        width = page_size$width,
        height = page_size$height),
      type = "continuous",
      page_margins = page_mar(
        bottom = page_margins$bottom,
        top = page_margins$top,
        right = page_margins$right,
        left = page_margins$left,
        header = page_margins$header,
        footer = page_margins$footer,
        gutter = page_margins$gutter)
    )
    defaut_sect_headers <- xml_find_all(docx_body_xml(x), "w:body/w:sectPr/w:headerReference")
    defaut_sect_headers <- lapply(defaut_sect_headers, as_xml_document)
    defaut_sect_footers <- xml_find_all(docx_body_xml(x), "w:body/w:sectPr/w:footerReference")
    defaut_sect_footers <- lapply(defaut_sect_footers, as_xml_document)

    x <- body_set_default_section(x, default_sect_properties)
    defaut_sect <- xml_find_first(docx_body_xml(x), "w:body/w:sectPr")

    for(i in rev(seq_len(length(defaut_sect_footers)))){
      xml_add_child(defaut_sect, defaut_sect_footers[[i]])
    }

    for(i in seq_len(length(defaut_sect_headers))){
      xml_add_child(defaut_sect, defaut_sect_headers[[i]])
    }

    forget(get_reference_rdocx)
    print(x, target = output_file)
    output_file
  }
  output_formats$bookdown_output_format = 'docx'
  output_formats

}
