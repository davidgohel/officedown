#' @importFrom utils modifyList
paragraph_settings <- function( align = "left", paddings = list(), borders = list(), shading = "transparent") {

  borders_ <- list(
    b = border_settings(width=0),
    l = border_settings(width=0),
    t = border_settings(width=0),
    r = border_settings(width=0) )
  borders_ <- modifyList(borders_, borders)

  paddings_ <- list(b = 3, t = 0, l = 0, r = 0)
  paddings_ <- modifyList(paddings_, paddings)

  obj = list()

  if( !is.null(shading) && !is.color( shading ) )
    stop("shading must be a valid color.", call. = FALSE )
  else obj$shading <- shading

  choices_align <- c("left", "right", "center", "justify")
  if( !valid_char(align) )
    stop("align must be a valid character scalar.", call. = FALSE )
  else if( ! align %in% choices_align )
    stop("align must be one of ", paste( shQuote(choices_align), collapse = ", "), call. = FALSE )
  obj$align <- align

  if( !valid_pos_num(paddings_$b) )
    stop("bottom padding must be a valid positive integer scalar.", call. = FALSE )
  else paddings_$b <- as.integer(paddings_$b)

  if( !valid_pos_num(paddings_$t) )
    stop("top padding must be a valid positive integer scalar.", call. = FALSE )
  else paddings_$t <- as.integer(paddings_$t)

  if( !valid_pos_num(paddings_$l) )
    stop("left padding must be a valid positive integer scalar.", call. = FALSE )
  else paddings_$l <- as.integer(paddings_$l)

  if( !valid_pos_num(paddings_$r) )
    stop("right padding must be a valid positive integer scalar.", call. = FALSE )
  else paddings_$r <- as.integer(paddings_$r)

  obj$borders <- borders_
  obj$paddings <- paddings_

  class( obj ) = "paragraph_settings"

  obj
}


format_paragraph_settings <- function(x, type = "docx", ...){

  os <- ""
  os <- paste0( os, "<w:par_settings>" )
  os <- paste0( os, sprintf("<w:jc w:val=\"%s\"/>", x$align ) )
  os <- paste0( os, "<w:pBdr>" )
  os <- paste0( os, format(x$borders$t, type = "docx", side = "top" ) )
  os <- paste0( os, format(x$borders$b, type = "docx", side = "bottom" ) )
  os <- paste0( os, format(x$borders$l, type = "docx", side = "left" ) )
  os <- paste0( os, format(x$borders$r, type = "docx", side = "right" ) )
  os <- paste0( os, "</w:pBdr>" )

  os <- paste0( os,
    sprintf("<w:spacing w:after=\"%.0f\" w:before=\"%.0f\"/>", x$paddings$b*20, x$paddings$t*20)
  )
  os <- paste0( os,
    sprintf("<w:ind w:left=\"%.0f\" w:right=\"%.0f\"/>", x$paddings$l*20, x$paddings$r*20)
  )

  if( !is.null(x$shading) && is_visible(x$shading) ){
    os <- paste0(os, sprintf("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>", to_rgb(x$shading)) )
  }
  os <- paste0( os, "</w:par_settings>" )
  os
}
