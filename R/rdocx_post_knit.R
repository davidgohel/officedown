#' @importFrom officer run_reference to_wml
as_reference <- function(z){
  str <- to_wml(run_reference(z))
  paste0("`<w:hyperlink w:anchor=\"", z, "\">", str, "</w:hyperlink>`{=openxml}")
}

post_knit_table_captions <- function(content, tab.cap.pre, tab.cap.sep, style) {
  is_captions <- grepl("<caption>\\(\\\\#tab:[-[:alnum:]]+\\)(.*)</caption>", content)
  if (any(is_captions)) {
    browser()
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

post_knit_references <- function(content, lp = ""){

  if(!lp %in% c("tab:", "fig:", "")){
    stop("lp must be one of `tab:`, `fig:` or empty ``.")
  }
  regexpr_str <- paste0('\\\\@ref\\(', lp, '([-[:alnum:]]+)\\)')

  gmatch <- gregexpr(regexpr_str, content)
  result <- regmatches(content,gmatch)

  result <- lapply(result, function(z){
    if(length(z) > 0){
      ids <- gsub(regexpr_str, '\\1', z)
      as_reference(ids)
    } else z
  })
  regmatches(content,gmatch) <- result
  content
}

