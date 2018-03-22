# macro ----
LIST_CHUNK_MACRO_PAR <- list(
  MACRO_PAGEBREAK = "chunk_page_break",
  MACRO_COLUMNBREAK = "chunk_column_break",
  MACRO_FTEXT = "macro_text_format",
  MACRO_PAR_SETTING = "macro_paragraph_format",
  MACRO_TEXT_STYLE = "macro_text_styled"
)

LIST_BLOCK_MACRO <- list(
  MACRO_TOC = "add_toc",
  MACRO_POUR_DOCX = "add_external_docx",
  MACRO_LANDSCAPE_START = "add_section_continuous",
  MACRO_LANDSCAPE_STOP = "add_section_landscape",
  MACRO_MULTICOL_START = "add_section_continuous",
  MACRO_MULTICOL_STOP = "add_section_columns"
)

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

chunk_to_ooxml <- function(text, fname, type) {
  text <- str_extract(text, paste0("(", yaml_pattern, "){0,1}[ ]*--->") )
  text <- gsub("[ ]*--->$", "", text)
  sapply( text, function(x, fname, type) {
    x <- yaml::yaml.load(x)
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
  for( i in names(LIST_CHUNK_MACRO_PAR) ){
    regex <- chunk_pattern(i)
    if( any( macro_ <- is_match(txt, regex) ) ){
      ooxml_str <- ooxml_values(txt[macro_], regex, LIST_CHUNK_MACRO_PAR[[i]], type = type)
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

