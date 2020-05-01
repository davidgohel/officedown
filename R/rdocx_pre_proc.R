# macro variables ----
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

# to ooxml ----
#' @importFrom yaml yaml.load
comment_tag_to_ooxml <- function(text, fname, type) {
  regex <- paste0("(", yaml_pattern, "){0,1}[ ]*--->")
  gmatch <- regexpr(regex, text)
  text <- regmatches(text, gmatch)
  text <- gsub("[ ]*--->$", "", text)

  # parse_arguments and call fname, for each call to_wml to get ooxml code
  sapply( text, function(x, fname, type) {
    x <- yaml.load(x)
    if( is.null(x) )
      x <- list()
    x <- do.call(fname, x)
    to_wml(x)
  }, fname = fname, type = type)
}

ooxml_values <- function(txt, regex, fname, type){
  gmatch <- gregexpr(regex, txt)
  all_extracts <- regmatches(txt, gmatch)
  ooxml_str <- lapply(all_extracts, comment_tag_to_ooxml, fname = fname, type = type)
  ooxml_str
}


block_macro <- function(txt, type = "docx"){
  for( i in names(LIST_BLOCK_MACRO) ){
    regex <- block_pattern(i)
    if( any( macro_ <- grepl(regex, txt) ) ){
      ooxml_str <- ooxml_values(txt[macro_], regex, LIST_BLOCK_MACRO[[i]], type = type)
      ooxml_str <- paste("```{=openxml}", ooxml_str, "```", sep = "\n")
      txt[macro_] <- ooxml_str
    }
  }
  txt
}
