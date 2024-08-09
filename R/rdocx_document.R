# utils ----
#' @importFrom utils getAnywhere getFromNamespace
get_fun <- function(x) {
  if (grepl("::", x, fixed = TRUE)) {
    coumpounds <- strsplit(x, split = "::", x, fixed = TRUE)[[1]]
    z <- getFromNamespace(coumpounds[2], ns = coumpounds[1])
  } else {
    z <- getAnywhere(x)
    if (length(z$objs) < 1) {
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

absolute_path <- function(x) {
  if (length(x) != 1L) {
    stop("'x' must be a single character string")
  }
  epath <- path.expand(x)
  if (file.exists(epath)) {
    epath <- normalizePath(epath, "/", mustWork = TRUE)
  } else {
    if (!dir.exists(dirname(epath))) {
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
    pre = "Table", sep = ":",
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
#' @description 'R Markdown' Format for converting from 'R Markdown'
#' document to an MS Word document.
#'
#' The function enhances the output offered by [rmarkdown::word_document()] with
#' advanced formatting features.
#' @param base_format a scalar character, the format to be used as a
#' base document for 'officedown'. Default to [word_document][rmarkdown::word_document] but
#' can also be `word_document2()` from bookdown.
#'
#' When the `base_format` used is `bookdown::word_document2`, the `number_sections`
#' parameter is automatically set to `FALSE`. Indeed, if you want numbered titles,
#' you are asked to use a Word document template with auto-numbered titles (the title
#' styles of the default `rdocx_document' template are already set to FALSE).
#'
#' @param tables see section 'Tables' below.
#' @param plots see section 'Plots' below.
#' @param lists see section 'Lists' below.
#' @param mapstyles a named list of style to be replaced in the generated
#' document. `list("Normal" = c("Author", "Date"))` will result in a document where
#' all paragraphs styled with stylename "Date" and "Author" will be then styled with
#' stylename "Normal".
#' @param md2 if TRUE sets number_section to true
#' @param reference_num if `TRUE`, text for references to sections will be
#' the section number (e.g. '3.2'). If FALSE, text for references to sections
#' will be the text (e.g. 'section title').
#' @param page_size,page_margins default page and margins dimensions. If
#' not null (the default), these values are used to define the default Word section.
#' See [officer::page_size()] and [officer::page_mar()].
#' @param ... arguments used by [word_document][rmarkdown::word_document]
#' @return R Markdown *output format* to pass to [render][rmarkdown::render].
#' @section Tables:
#'
#' ```{r child = "man/rdocx/tables.Rmd"}
#' ```
#' @section Plots:
#'
#' ```{r child = "man/rdocx/plots.Rmd"}
#' ```
#' @section Lists:
#'
#' ```{r child = "man/rdocx/lists.Rmd"}
#' ```
#' @section Finding stylenames:
#'
#' ```{r child = "man/rdocx/stylenames.Rmd"}
#' ```
#'
#' @section R Markdown yaml:
#'
#' ```{r child = "man/rdocx/rmarkdown-yaml.Rmd"}
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
#' # rdocx_document basic example -----
#' @example examples/rdocx_document.R
#' @importFrom officer change_styles block_pour_docx
#' @importFrom utils modifyList
rdocx_document <- function(base_format = "rmarkdown::word_document",
                           tables = list(), plots = list(), lists = list(),
                           mapstyles = list(), page_size = list(), page_margins = list(),
                           reference_num = TRUE, md2 = TRUE, ...) {
  args <- list(...)
  if (is.null(args$reference_docx)) {
    args$reference_docx <- system.file(
      package = "officedown", "examples",
      "bookdown", "template.docx"
    )
  }
  if (!is.null(args$number_sections) && isTRUE(args$number_sections)) {
    if (isTRUE(md2)) {
      args$number_sections <- TRUE
    } else {
      args$number_sections <- FALSE
    }
  }

  args$reference_docx <- absolute_path(args$reference_docx)

  base_format_fun <- get_fun(base_format)
  output_formats <- do.call(base_format_fun, args)


  tables <- modifyList(tables_default_values, tables)
  plots <- modifyList(plots_default_values, plots)
  lists <- modifyList(lists_default_values, lists)

  if (!is.null(page_size)) {
    page_size <- modifyList(page_size_default_values, page_size)
  }
  if (!is.null(page_margins)) {
    page_margins <- modifyList(page_mar_default_values, page_margins)
  }

  output_formats$knitr$opts_chunk <- append(
    output_formats$knitr$opts_chunk,
    list(
      tab.cap.style = tables$caption$style,
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
      fig.topcaption = plots$topcaption,
      is_rdocx_document = TRUE
    )
  )
  if (is.null(output_formats$knitr$knit_hooks)) {
    output_formats$knitr$knit_hooks <- list()
  }
  output_formats$knitr$knit_hooks$plot <- plot_word_fig_caption

  # This is the intermediate_dir and wll be updated if `intermediates_generator` is called
  intermediate_dir <- "."

  temp_intermediates_generator <- output_formats$intermediates_generator
  output_formats$intermediates_generator <- function(...) {
    intermediate_dir <<- list(...)[[2]]
    temp_intermediates_generator(...)
  }

  output_formats$post_knit <- function(metadata, input_file, runtime, ...) {
    output_file <- file_with_meta_ext(input_file, "knit", "md")
    output_file <- file.path(intermediate_dir, output_file)
    if (!file.exists(output_file)) {
      output_file <- gsub("\\.spin\\.Rmd$", ".knit.md", input_file)
      output_file <- file.path(intermediate_dir, output_file)
    }
    if (!file.exists(output_file)) {
      stop("can not find md file for 'post-knit' treatment.")
    }

    content <- readLines(output_file)
    content <- post_knit_caption_references(content, lp = tables$tab.lp)
    content <- post_knit_caption_references(content, lp = plots$fig.lp)
    content <- post_knit_std_references(content, numbered = reference_num)
    content <- block_macro(content)
    writeLines(content, output_file)
  }


  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_docx(output_file)
    x <- process_par_settings(x)
    x <- process_list_settings(x, ul_style = lists$ul.style, ol_style = lists$ol.style)
    x <- change_styles(x, mapstyles = mapstyles)

    if (!is.null(page_size) && !is.null(page_margins)) {
      # default section
      default_sect_properties <- prop_section(
        page_size = page_size(
          orient = page_size$orient,
          width = page_size$width,
          height = page_size$height
        ),
        type = "continuous",
        page_margins = page_mar(
          bottom = page_margins$bottom,
          top = page_margins$top,
          right = page_margins$right,
          left = page_margins$left,
          header = page_margins$header,
          footer = page_margins$footer,
          gutter = page_margins$gutter
        )
      )
      x <- body_set_default_section(x, default_sect_properties)
    }

    forget(get_reference_rdocx)
    print(x, target = output_file)
    output_file
  }
  output_formats$bookdown_output_format <- "docx"
  output_formats
}
