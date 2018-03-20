
# sections ----
#' @export
macro_section_continuous <- function( ){
  str <- "<w:p><w:pPr><w:sectPr><w:worded/><w:type w:val=\"continuous\"/></w:sectPr></w:pPr></w:p>"
  class(str) <- "ooxml"
  str
}

#' @export
macro_section_landscape <- function( w = 21 / 2.54, h = 29.7 / 2.54 ){
  w = w * 20 * 72
  h = h * 20 * 72
  pgsz_str <- "<w:pgSz w:orient=\"landscape\" w:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, h, w )
  str <- sprintf( "<w:p><w:pPr><w:sectPr><w:worded/>%s</w:pPr></w:p>", pgsz_str)
  class(str) <- "ooxml"
  str
}

#' @export
macro_section_portrait <- function( w = 21 / 2.54, h = 29.7 / 2.54 ){
  w = w * 20 * 72
  h = h * 20 * 72
  pgsz_str <- "<w:pgSz w:orient=\"portrait\" w:w=\"%.0f\" w:h=\"%.0f\"/>"
  pgsz_str <- sprintf(pgsz_str, w, h )
  str <- sprintf( "<w:p><w:pPr><w:sectPr><w:worded/>%s</w:pPr></w:p>", pgsz_str)
  class(str) <- "ooxml"
  str
}

#' @export
macro_section_columns <- function(widths = c(2.5,2.5), space = .25, sep = FALSE){

  widths <- widths * 20 * 72
  space <- space * 20 * 72

  columns_str_all_but_last <- sprintf("<w:col w:w=\"%.0f\" w:space=\"%.0f\"/>",
                                      widths[-length(widths)], space)
  columns_str_last <- sprintf("<w:col w:w=\"%.0f\"/>",
                              widths[length(widths)])
  columns_str <- c(columns_str_all_but_last, columns_str_last)

  if( length(widths) < 2 )
    stop("length of widths should be at least 2", call. = FALSE)

  columns_str <- sprintf("<w:cols w:num=\"%.0f\" w:sep=\"%.0f\" w:space=\"%.0f\" w:equalWidth=\"0\">%s</w:cols>",
                         length(widths), as.integer(sep), space, paste0(columns_str, collapse = "") )

  str <- paste0( "<w:p>",
                 "<w:pPr><w:sectPr><w:worded/>",
                 "<w:type w:val=\"continuous\"/>",
                 columns_str, "</w:sectPr></w:pPr></w:p>")
  class(str) <- "ooxml"
  str
}

#' @export
chunk_column_break <- function( ){
  str <- "<w:r><w:br w:type=\"column\"/></w:r>"
  class(str) <- "ooxml_chunk"
  str
}


# misc ----
#' @export
block_docx_pour <- function(docx_file){
  str <- paste0("<w:altChunk r:id=\"", docx_file, "\"/>")
  class(str) <- "ooxml"
  str
}

#' @export
block_toc <- function( level = 3, style = NULL, separator = ";"){

  if( is.null( style )){
    str <- paste0("<w:p>", "<w:pPr/>",
                  "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>",
                  "<w:r><w:instrText xml:space=\"preserve\" w:dirty=\"true\">TOC \u005Co &quot;1-%.0f&quot; \u005Ch \u005Cz \u005Cu</w:instrText></w:r>",
                  "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                  "</w:p>")
    str <- sprintf(str, level)
  } else {
    str <- paste0("<w:p>", "<w:pPr/>",
                  "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>",
                  "<w:r><w:instrText xml:space=\"preserve\" w:dirty=\"true\">TOC \u005Ch \u005Cz \u005Ct \"%s%s1\"</w:instrText></w:r>",
                  "<w:r><w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/></w:r>",
                  "</w:p>")
    str <- sprintf(str, style, separator)
  }

  class(str) <- "ooxml"
  str
}

#' @export
macro_paragraph_format <- function( text.align = "left",
                                    border = fp_border(width=0),
                                    padding.bottom = 3, padding.top = 0,
                                    padding.left = 0, padding.right = 0,
                                    border.bottom = borderfp(width=0),
                                    border.left = borderfp(width=0),
                                    border.top = borderfp(width=0),
                                    border.right = borderfp(width=0),
                                    shading.color = "transparent" ){

  x <- parfp(text.align = text.align,
        padding.bottom = padding.bottom, padding.top = padding.top,
        padding.left = padding.left, padding.right = padding.right,
        border.bottom = border.bottom, border.left = border.left,
        border.top = border.top, border.right = border.right,
        shading.color = shading.color)
  str <- format_parfp(x, type = "docx")
  class(str) <- "ooxml_chunk"
  str
}

#' @export
chunk_page_break <- function( ){

  str <- "<w:r><w:br w:type=\"page\"/></w:r>"
  class(str) <- "ooxml_chunk"
  str
}

#' @export
chunk_text <- function( str, style ){
  str <- paste0("<w:r>",
                sprintf("<w:rPr><w:rStyle w:chunk_style=\"%s\"/></w:rPr>", style),
                sprintf("<w:t xml:space=\"preserve\">%s</w:t>", str),
                "</w:r>")
  class(str) <- "ooxml_chunk"
  str
}

#' @export
chunk_ftext <- function( str, fp ){

  str <- paste0("<w:r>",
                format(fp, type = "docx"),
                sprintf("<w:t xml:space=\"preserve\">%s</w:t>", str),
                "</w:r>")
  class(str) <- "ooxml_chunk"
  str
}


# seqfield ----
#' @export
chunk_seqfield <- function( str, style ){
  xml_elt_1 <- paste0("<w:r>",
                      sprintf("<w:rPr><w:rStyle w:chunk_style=\"%s\"/></w:rPr>", style),
                      "<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/>",
                      "</w:r>")
  xml_elt_2 <- paste0("<w:r>",
                      sprintf("<w:rPr><w:rStyle w:chunk_style=\"%s\"/></w:rPr>", style),
                      sprintf("<w:instrText xml:space=\"preserve\" w:dirty=\"true\">%s</w:instrText>", str ),
                      "</w:r>")
  xml_elt_3 <- paste0("<w:r>",
                      sprintf("<w:rPr><w:rStyle w:chunk_style=\"%s\"/></w:rPr>", style),
                      "<w:fldChar w:fldCharType=\"end\" w:dirty=\"true\"/>",
                      "</w:r>")
  str <- paste0(xml_elt_1, xml_elt_2, xml_elt_3)
  class(str) <- "ooxml_chunk"
  str
}
