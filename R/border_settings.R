border_styles = c( "none", "solid", "dotted", "dashed" )

border_settings = function( color = "black", style = "solid", width = 1 ){

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

  class( out ) = "border_settings"
  out
}

format.border_settings = function (x, side = "left", type = "docx", ...){

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
