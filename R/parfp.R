parfp = function(text.align = "left",
                 padding.bottom = 3, padding.top = 0,
                 padding.left = 0, padding.right = 0,
                 border.bottom = borderfp(width=0),
                 border.left = borderfp(width=0),
                 border.top = borderfp(width=0),
                 border.right = borderfp(width=0),
                 shading.color = "transparent") {

  obj = list()

  if( !is.null(shading.color) && !is.color( shading.color ) )
    stop("shading.color must be a valid color.", call. = FALSE )
  else obj$shading.color <- shading.color

  choices_align <- c("left", "right", "center", "justify")
  if( !valid_char(text.align) )
    stop("text.align must be a valid character scalar.", call. = FALSE )
  else if( ! text.align %in% choices_align )
    stop("text.align must be one of ", paste( shQuote(choices_align), collapse = ", "), call. = FALSE )
  obj$text.align <- text.align

  if( !valid_pos_num(padding.bottom) )
    stop("padding.bottom must be a valid positive integer scalar.", call. = FALSE )
  else obj$padding.bottom <- as.integer(padding.bottom)

  if( !valid_pos_num(padding.top) )
    stop("padding.top must be a valid positive integer scalar.", call. = FALSE )
  else obj$padding.top <- as.integer(padding.top)

  if( !valid_pos_num(padding.left) )
    stop("padding.left must be a valid positive integer scalar.", call. = FALSE )
  else obj$padding.left <- as.integer(padding.left)

  if( !valid_pos_num(padding.right) )
    stop("padding.right must be a valid positive integer scalar.", call. = FALSE )
  else obj$padding.right <- as.integer(padding.right)

  obj$border.bottom <- border.bottom
  obj$border.top <- border.top
  obj$border.left <- border.left
  obj$border.right <- border.right

  class( obj ) = "parfp"

  obj
}


format_parfp <- function(x, type = "docx", ...){

  os <- ""
  os <- paste0( os, "<w:pPrworded>" )
  os <- paste0( os, sprintf("<w:jc w:val=\"%s\"/>", x$text.align ) )
  os <- paste0( os, "<w:pBdr>" )
  os <- paste0( os, format(x$border.top, type = "docx", side = "top" ) )
  os <- paste0( os, format(x$border.bottom, type = "docx", side = "bottom" ) )
  os <- paste0( os, format(x$border.left, type = "docx", side = "left" ) )
  os <- paste0( os, format(x$border.right, type = "docx", side = "right" ) )
  os <- paste0( os, "</w:pBdr>" )

  os <- paste0(
    os,
    sprintf("<w:spacing w:after=\"%.0f\" w:before=\"%.0f\"/>", x$padding.bottom*20, x$padding.top*20)
  )
  os <- paste0(
    os,
    sprintf("<w:ind w:left=\"%.0f\" w:right=\"%.0f\"/>", x$padding.left*20, x$padding.right*20)
  )

  if( !is.null(x$shading.color) && is_visible(x$shading.color) ){
    os <- paste0(os, sprintf("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"%s\"/>", to_rgb(x$shading.color)) )
  }
  os <- paste0( os, "</w:pPrworded>" )
  os
}
