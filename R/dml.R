#' Make A DrawingML plotting object to print in Powerpoint
#'
#' Only Powerpoint outputs currently supported. Pandoc >=2.4 is required.
#' Currently plots with rasters are not supported.
#'
#' @param code plotting instructions
#' @param ggobj ggplot object to print. argument code will be ignored if this
#'   argument is supplied.
#' @param layout either a character indicating layout type: "default",
#'   "left_col", "right_col", "full_slide", or a list with elements "width",
#'   "height", "left", and "top" in inches. Layouts are drawn from the
#'   powerpoint template used for the R Markdown document.
#' @param bg,fonts,pointsize,editable Parameters passed to \code{\link{dml_pptx}}
#' @export
#' @importFrom rvg dml_pptx
dml <- function(code, ggobj = NULL, layout = "default",
                bg = "white", fonts = list(), pointsize = 12, editable = TRUE) {
  out <- list()
  out$code <- substitute(code)
  out$ggobj <- ggobj
  if (!(all(layout %in% c("default", "left_col", "right_col", "full_slide")) ||
        (inherits(layout, "list") &&
         setequal(names(layout), c("width", "height", "left", "top"))))) {
    stop("Invalid layout value: ", names(layout))
  }
  out$layout <- layout
  out$bg <- bg
  out$fonts <- fonts
  out$pointsize <- pointsize
  out$editable <- editable
  class(out) <- "dml"
  return(out)
}

#' @title Render a plot as a Powerpint DrawingML object
#' @description Function used to render DrawingML in knitr/rmarkdown documents.
#' Only Powerpoint outputs currently supported
#'
#' @param x a \code{dml} object
#' @param ... further arguments, not used.
#' @author Noam Ross
#' @importFrom knitr knit_print asis_output opts_knit
#' @importFrom rmarkdown pandoc_version
#' @importFrom xml2 xml_find_first
#' @importFrom grDevices dev.off
#' @export
knit_print.dml <- function(x, ...) {
  if (pandoc_version() < 2.4) {
    stop("pandoc version >= 2.4 required for DrawingML output in pptx")
  }

  if (is.null(opts_knit$get("rmarkdown.pandoc.to")) ||
      opts_knit$get("rmarkdown.pandoc.to") != "pptx") {
    stop("DrawingML currently only supported for pptx output")
  }

  if (inherits(x$layout, "list")) {
    id_xfrm <- x$layout
  } else {
    id_xfrm <- get_content_layout(x$layout)
  }

  dml_file <- tempfile(fileext = ".dml")
  img_directory = get_img_dir()

  dml_pptx(file = dml_file, width = id_xfrm$width, height = id_xfrm$height,
           offx = id_xfrm$left, offy = id_xfrm$top, pointsize = x$pointsize,
           last_rel_id = 1L,
           editable = x$editable, standalone = FALSE, raster_prefix = img_directory)

  tryCatch({
    if (!is.null(x$ggobj) ) {
      stopifnot(inherits(x$ggobj, "ggplot"))
      print(x$ggobj)
    } else {
      eval(x$code)
    }
  }, finally = dev.off() )

  dml_xml <- read_xml(dml_file)
  raster_files <- list_raster_files(img_dir = img_directory )

  if (length(raster_files)) {
    rast_element <- xml_find_all(dml_xml, "//p:pic/p:blipFill/a:blip")
    raster_files <- list_raster_files(img_dir = img_directory )
    raster_id <- xml_attr(rast_element, "embed")
    for (i in seq_along(raster_files)) {
      xml_attr(rast_element[i], "r:embed") <- raster_files[i]
    }
  }
  dml_str <- paste(
    as.character(xml_find_first(dml_xml, "//p:grpSp")),
    collapse = "\n"
  )

  knit_print(asis_output(
    x = paste("```{=openxml}", dml_str, "```", sep = "\n")
  ))
}



#' If size is not provided, get the size of the main content area of the slide
#' @noRd
get_content_layout_uncached <- function(layout) {
  ref_pptx <- read_pptx(get_reference_pptx())

  if (layout == "full_slide") {
    ph_location_fullsize(x = ref_pptx)
  } else if (layout == "default") {
    ph_location_type(layout = "Title and Content",
                     master = "Office Theme", type = "body", x = ref_pptx)
  } else if (layout == "left_col") {
    ph_location_left(x = ref_pptx)
  } else if (layout == "right_col") {
    ph_location_right(x = ref_pptx)
  } else {
    stop("Unknown layout type")
  }
}

#' @importFrom memoise memoise
#' @noRd
get_content_layout <- memoise(get_content_layout_uncached)


#' @export
print.dml <- function(x, ...) {
  cat("DrawingML object with parameters:\n")
  cat(paste0("  ", names(x), ": ", as.character(x), collapse = "\n"))
}

get_img_dir <- function(){
  uid <- basename(tempfile(pattern = ""))
  img_directory = file.path(getwd(), uid )
  img_directory
}
list_raster_files <- function(img_dir){
  path_ <- dirname(img_dir)
  uid <- basename(img_dir)
  list.files(path = path_, pattern = paste0("^", uid, "(.*)\\.png$"), full.names = TRUE )
}

