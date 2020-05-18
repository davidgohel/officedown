#' @importFrom officer run_reference to_wml
as_reference_caption <- function(z){
  sf <- run_seqfield(seqfield = paste0(" REF ", z, " \\h"))
  str <- to_wml(sf)
  paste0("`<w:hyperlink w:anchor=\"", z, "\">", str, "</w:hyperlink>`{=openxml}")
}
as_reference_standard <- function(z, numbered = FALSE){
  str <- paste0(" REF ", z, " \\h")
  if(numbered)
    str <- paste0(str, " \\r")

  sf <- run_seqfield(seqfield = str)
  str <- to_wml(sf)
  paste0("`<w:hyperlink w:anchor=\"", z, "\">", str, "</w:hyperlink>`{=openxml}")
}

post_knit_table_captions <- function(content, tab.cap.pre, tab.cap.sep, style) {
  is_captions <- grepl("<caption>\\(\\\\#tab:[-[:alnum:]]+\\)(.*)</caption>", content)
  if (any(is_captions)) {
    captions <- content[is_captions]
    ids <- gsub("<caption>\\(\\\\#tab:([-[:alnum:]]+)\\)(.*)</caption>", "\\1", captions)
    labels <- gsub("<caption>\\(\\\\#tab:[-[:alnum:]]+\\)(.*)</caption>", "\\1", captions)
    str <- mapply(function(label, id) {
      pandoc_wml_caption(
        cap = label, cap.style = style,
        cap.pre = tab.cap.pre, cap.sep = tab.cap.sep,
        id = id, seq_id = "tab:"
      )
    }, label = labels, id = ids, SIMPLIFY = FALSE)
    content[is_captions] <- unlist(str)
  }
  content
}

post_knit_caption_references <- function(content, lp = ""){

  if(!lp %in% c("tab:", "fig:")){
    stop("lp must be one of `tab:`, `fig:`.")
  }
  regexpr_str <- paste0('\\\\@ref\\(', lp, '([-[:alnum:]]+)\\)')

  gmatch <- gregexpr(regexpr_str, content)
  result <- regmatches(content,gmatch)

  result <- lapply(result, function(z){
    if(length(z) > 0){
      ids <- gsub(regexpr_str, '\\1', z)
      as_reference_caption(ids)
    } else z
  })
  regmatches(content,gmatch) <- result
  content
}

post_knit_std_references <- function(content, numbered = FALSE){

  regexpr_str <- paste0('\\\\@ref\\(([-[:alnum:]]+)\\)')

  gmatch <- gregexpr(regexpr_str, content)
  result <- regmatches(content,gmatch)

  result <- lapply(result, function(z){
    if(length(z) > 0){
      ids <- gsub(regexpr_str, '\\1', z)
      as_reference_standard(ids, numbered = numbered)
    } else z
  })
  regmatches(content,gmatch) <- result
  content
}

