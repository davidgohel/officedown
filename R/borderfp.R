border_styles = c( "none", "solid", "dotted", "dashed" )

#' @title border properties object
#'
#' @description create a border properties object.
#'
#' @param color border color - single character value (e.g. "#000000" or "black")
#' @param style border style - single character value : "none" or "solid" or "dotted" or "dashed"
#' @param width border width - an integer value : 0>= value
#' @examples
#' borderfp()
#' borderfp(color="orange", style="solid", width=1)
#' borderfp(color="gray", style="dotted", width=1)
#' @export
borderfp = function( color = "black", style = "solid", width = 1 ){

  out <- list()

  if( !valid_pos_num(width) )
    stop("width must be a valid positive integer scalar.", call. = FALSE )
  else out$width <- as.integer(width)

  if( !is.null(color) && !is.color( color ) )
    stop("color must be a valid color.", call. = FALSE )
  else out$color <- color

  if( !valid_char(style) )
    stop("style must be a valid character scalar.", call. = FALSE )
  else if( ! style %in% border_styles )
    stop("style must be one of ", paste( shQuote(border_styles), collapse = ", "), call. = FALSE )
  out$style <- style

  class( out ) = "borderfp"
  out
}

#' @export
#' @rdname borderfp
#' @param x \code{borderfp} object
#' @param type output type
format.borderfp = function (x, type = "docx", side = "left", ...){

  if( !is_visible(x$color) || x$width < .001 || x$style == "none") {
    return("")
  }
  if( x$style == "solid")
    type <- "single"
  else if( x$style == "dotted")
    type <- "dotted"
  else if( x$style == "dashed")
    type <- "dashed"

  sprintf("<w:%s w:val=\"%s\" w:sz=\"%.0f\" w:space=\"0\" w:color=\"%s\"/>",
          side, type, x$width * 8, to_rgb(x$color) )
}
