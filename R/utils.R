htmlEscapeCopy <- local({

  .htmlSpecials <- list(
    `&` = '&amp;',
    `<` = '&lt;',
    `>` = '&gt;'
  )
  .htmlSpecialsPattern <- paste(names(.htmlSpecials), collapse='|')
  .htmlSpecialsAttrib <- c(
    .htmlSpecials,
    `'` = '&#39;',
    `"` = '&quot;',
    `\r` = '&#13;',
    `\n` = '&#10;'
  )
  .htmlSpecialsPatternAttrib <- paste(names(.htmlSpecialsAttrib), collapse='|')
  function(text, attribute=FALSE) {
    pattern <- if(attribute)
      .htmlSpecialsPatternAttrib
    else
      .htmlSpecialsPattern
    text <- enc2utf8(as.character(text))
    # Short circuit in the common case that there's nothing to escape
    if (!any(grepl(pattern, text, useBytes = TRUE)))
      return(text)
    specials <- if(attribute)
      .htmlSpecialsAttrib
    else
      .htmlSpecials
    for (chr in names(specials)) {
      text <- gsub(chr, specials[[chr]], text, fixed = TRUE, useBytes = TRUE)
    }
    Encoding(text) <- "UTF-8"
    return(text)
  }
})


merge_pPr <- function(new, current, xpath){
  jc <- xml_child(new, xpath)
  jc_ref <- xml_child(current, xpath)
  if(inherits(jc, "xml_missing")) return(FALSE)
  if( !inherits(jc_ref, "xml_missing") )
    xml_replace(jc_ref, jc)
  else xml_add_child(current, jc)

  TRUE

}

#' @importFrom uuid UUIDgenerate
as_bookmark_md <- function(id, str) {
  new_id <- uuid::UUIDgenerate()
  bm_start_str <- sprintf("`<w:bookmarkStart w:id=\"%s\" w:name=\"%s\"/>`{=openxml}", new_id, id)
  bm_start_end <- sprintf("`<w:bookmarkEnd w:id=\"%s\"/>`{=openxml}", new_id)
  paste0(bm_start_str, str, bm_start_end)
}

pandoc_wml_caption <- function(cap = NULL, cap.style = NULL, cap.pre = NULL, cap.sep = NULL, id = NULL, seq_id = NULL,
                               tnd = 0, tns = "-", prop = NULL,
                               ...){

  if( is.null(cap)) return("")

  run_str <- cap
  if(!is.null(seq_id)){
    seq_id <- gsub(":$", "", seq_id)
    autonum <- run_autonum(seq_id = seq_id,
                           pre_label = cap.pre,
                           post_label = cap.sep, bkm = id, bkm_all = FALSE,
                           tnd = tnd, tns = tns, prop = prop)
    autonum <- paste("`", to_wml(autonum), "`{=openxml}", sep = "")
    run_str <- paste0(autonum, run_str)
  }

  paste0(if(!is.null(cap.style)) paste0("\n\n::: {custom-style=\"", cap.style, "\"}"),
         "\n\n",
         "<caption>",
         run_str,
         "</caption>",
         if(!is.null(cap.style)) paste0("\n:::\n"),
         "\n\n")

}


str_lorem <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla eu pulvinar arcu, quis aliquam dui. In at cursus ante. Vestibulum non sagittis lacus. Duis vitae iaculis dui. Vivamus tempor, nibh ut pretium tempus, enim lorem dignissim quam, at euismod massa magna at magna. Sed facilisis dapibus diam nec volutpat. Maecenas facilisis dapibus egestas. Curabitur dignissim pharetra pulvinar. Nunc bibendum elit sed cursus congue. Curabitur ligula quam, iaculis faucibus orci quis, vestibulum lobortis lectus. Suspendisse fringilla nisl pulvinar, laoreet tellus sed, sollicitudin tortor. Donec consequat congue erat in iaculis. Curabitur luctus tellus ut turpis iaculis, nec laoreet ligula scelerisque."

