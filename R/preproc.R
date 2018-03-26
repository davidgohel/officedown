# macro variables ----
LIST_CHUNK_MACRO <- list(
  CHUNK_PAGEBREAK = "chunk_page_break",
  CHUNK_COLUMNBREAK = "chunk_column_break",
  CHUNK_TEXT = "chunk_styled_text",
  CHUNK_TEXT_STYLE = "chunk_text_stylenamed"
)

LIST_BLOCK_MACRO <- list(
  BLOCK_TOC = "block_toc",
  BLOCK_POUR_DOCX = "block_pour_docx",
  BLOCK_LANDSCAPE_START = "block_section_continuous",
  BLOCK_LANDSCAPE_STOP = "block_section_landscape",
  BLOCK_MULTICOL_START = "block_section_continuous",
  BLOCK_MULTICOL_STOP = "block_section_columns"
)

# macro regexpr ----
yaml_pattern <- "\\{[^\\}]+\\}"

chunk_pattern <- function( id ){
  paste0("(<!---)([ ]*", id, "[ ]*)(", yaml_pattern, "){0,1}([ ]*--->)")
}
block_pattern <- function( id ){
  paste0("^([ ]*)", chunk_pattern(id), "([ ]*)$")
}

is_match <- function(str, regex){
  grepl(regex, str)
}

# to ooxml ----
#' @importFrom yaml yaml.load
chunk_to_ooxml <- function(text, fname, type) {
  text <- str_extract(text, paste0("(", yaml_pattern, "){0,1}[ ]*--->") )
  text <- gsub("[ ]*--->$", "", text)
  sapply( text, function(x, fname, type) {
    x <- yaml.load(x)
    if( is.null(x) )
      x <- list()
    x <- do.call(fname, x)
    format(x, type = type)
  }, fname = fname, type = type)
}

ooxml_values <- function(txt, regex, fname, type){
  all_extracts <- str_extract_all(txt, regex )
  ooxml_str <- lapply(all_extracts, chunk_to_ooxml, fname = fname, type = type)
  ooxml_str
}

#' @import stringr
chunk_macro <- function(txt, type = "docx"){
  for( i in names(LIST_CHUNK_MACRO) ){
    regex <- chunk_pattern(i)
    if( any( macro_ <- is_match(txt, regex) ) ){
      ooxml_str <- ooxml_values(txt[macro_], regex, LIST_CHUNK_MACRO[[i]], type = type)
      new_txt <- mapply(function(txtline, replacts){
        chunks <- c(txtline, replacts)
        chunks <- chunks[order(c(seq_along(txtline) * 2 - 1, seq_along(replacts) * 2))]
        do.call(paste0, as.list(chunks))
      }, str_split(txt[macro_], pattern = regex), ooxml_str, SIMPLIFY = FALSE)
      txt[macro_] <- as.character(unlist(new_txt))
    }
  }
  txt
}


block_macro <- function(txt, type = "docx"){
  for( i in names(LIST_BLOCK_MACRO) ){
    regex <- block_pattern(i)
    if( any( macro_ <- is_match(txt, regex) ) ){
      ooxml_str <- ooxml_values(txt[macro_], regex, LIST_BLOCK_MACRO[[i]], type = type)
      txt[macro_] <- as.character(ooxml_str)
    }
  }
  txt
}
