#' @title Render a plot as a Powerpint DrawingML object
#' @description Function used to render DrawingML in knitr/rmarkdown documents.
#' Only Powerpoint outputs currently supported
#'
#' @param x a `dml` object
#' @param ... further arguments, not used.
#' @author Noam Ross
#' @importFrom knitr knit_print asis_output opts_knit opts_current
#' @importFrom rmarkdown pandoc_version
#' @importFrom rvg dml dml_pptx
#' @importFrom grDevices dev.off
#' @importFrom rlang eval_tidy
#' @noRd
knit_print.dml <- function(x, ...) {
  if (pandoc_version() < 2.4) {
    stop("pandoc version >= 2.4 required for DrawingML output in pptx")
  }

  if (is.null(opts_knit$get("rmarkdown.pandoc.to")) ||
      opts_knit$get("rmarkdown.pandoc.to") != "pptx") {
    stop("DrawingML currently only supported for pptx output")
  }

  layout <- knitr::opts_current$get("layout")
  master <- knitr::opts_current$get("master")
  doc <- get_reference_pptx()
  if(is.null( ph <- knitr::opts_current$get("ph") )){
    ph <- officer::ph_location_type(type = "body")
  }
  if(!inherits(ph, "location_str")){
    stop("ph should be a placeholder location; ",
         "see officer::placeholder location for an example.",
         call. = FALSE)
  }

  id_xfrm <- get_content_ph(ph, layout, master, doc)

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
      rlang::eval_tidy(x$code)
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
#' @importFrom officer fortify_location
get_ph_uncached <- function(ph, layout, master, doc) {
  ls <- layout_summary(doc)

  if(!master %in% ls$master){
    stop("could not find master ", master, call. = FALSE)
  }

  slide_index <- which(ls$layout %in% layout & ls$master %in% master)
  if(length(slide_index)<1){
    stop("could not find layout ", layout, " and master ", master, call. = FALSE)
  }

  doc <- on_slide(doc, index = slide_index)

  fortify_location(ph, doc)
}

#' @importFrom memoise memoise
#' @noRd
get_content_ph <- memoise(get_ph_uncached)



get_img_dir <- function(){
  uid <- basename(tempfile(pattern = ""))
  img_directory = file.path(tempdir(), uid )
  img_directory
}
list_raster_files <- function(img_dir){
  path_ <- dirname(img_dir)
  uid <- basename(img_dir)
  list.files(path = path_, pattern = paste0("^", uid, "(.*)\\.png$"), full.names = TRUE )
}

