#' @import officer xml2

# list processing ----

pandoc_abstract_num_id <- function(doc, numbering_doc){

  out <- list(ul = NULL, ol = NULL)

  num_id <- xml_find_all(doc, "w:body/w:p/w:pPr/w:numPr/w:numId")
  num_id <- unique(xml_attr(num_id, "val"))

  if(length(num_id) < 1) {
    return(out)
  }

  xpath <- paste0(
    "w:num[",
    paste0(sprintf("contains(@w:numId, '%s')", num_id), collapse = " or "),
    "]/w:abstractNumId"
  )
  abstract_num <- xml_find_all(numbering_doc, xpath)
  abstract_num <- unique(xml_attr(abstract_num, "val"))
  abstract_num

  unordered_xpath <- paste0(
    "w:abstractNum[",
    "(",
    paste0(sprintf("contains(@w:abstractNumId, '%s')", abstract_num), collapse = " or "),
    ") and (",
    "w:lvl[position()<2]/w:numFmt/@w:val='bullet'",
    ")",
    "]"
  )
  unordered_abstract_num_id <- xml_find_first(numbering_doc, unordered_xpath)
  unordered_abstract_num_id <- xml_attr(unordered_abstract_num_id, "abstractNumId")

  ordered_xpath <- paste0(
    "w:abstractNum[",
    "(",
    paste0(sprintf("contains(@w:abstractNumId, '%s')", abstract_num), collapse = " or "),
    ") and (",
    "w:lvl[position()<2]/w:numFmt/@w:val='decimal'",
    ")",
    "]"
  )
  ordered_abstract_num_id <- xml_find_first(numbering_doc, ordered_xpath)
  ordered_abstract_num_id <- xml_attr(ordered_abstract_num_id, "abstractNumId")

  list(ul = unordered_abstract_num_id, ol = ordered_abstract_num_id)
}

replace_list_style <- function(style_id, abstract_num_id, numbering_doc){
  xpath <- sprintf("w:abstractNum[w:styleLink[contains(@w:val, '%s')]]", style_id)
  abstractnum_node <- xml_find_first(numbering_doc, xpath)
  new_abstract_num_id <- xml_attr(abstractnum_node, "abstractNumId")

  xpath <- paste0(
    "w:num/",
    "w:abstractNumId",
    sprintf("[contains(@w:val, '%s')]", abstract_num_id)
  )
  abstract_num <- xml_find_all(numbering_doc, xpath)
  xml_attr(abstract_num, "w:val") <- new_abstract_num_id
  TRUE
}

process_list_settings <- function(rdoc, ul_style = NULL, ol_style = NULL){

  numbering_file <- file.path(rdoc$package_dir, "word/numbering.xml")
  numbering_doc <- read_xml(numbering_file)

  if(is.null(ul_style) && is.null(ol_style)){
    return(rdoc)
  }

  si <- styles_info(rdoc, type = "numbering")

  # find pandoc ol and ul abstract_num_id
  abstract_num <- pandoc_abstract_num_id(
    doc = docx_body_xml(rdoc),
    numbering_doc = numbering_doc)
  if(!is.null(abstract_num$ul) && !is.null(ul_style)){
    ul_style_id <- si[si$style_name %in% ul_style, "style_id"]
    if(length(ul_style_id)){
      replace_list_style(
        style_id = ul_style_id,
        abstract_num_id = abstract_num$ul,
        numbering_doc = numbering_doc)
    } else {
      warning("numbering style ", shQuote(ul_style),
              " has not been found in the reference_docx document.",
              call. = FALSE)
    }
  }


  if(!is.null(abstract_num$ol) && !is.null(ol_style)){

    ol_style_id <- si[si$style_name %in% ol_style, "style_id"]
    if(length(ol_style_id)){
      replace_list_style(
        style_id = ol_style_id,
        abstract_num_id = abstract_num$ol,
        numbering_doc = numbering_doc)
    } else {
      warning("numbering style ", shQuote(ol_style),
              " has not been found in the reference_docx document.",
              call. = FALSE)
    }
  }
  write_xml(x = numbering_doc, file = numbering_file)
  rdoc
}

# ooxml reprocess with rdocx available ----

#' @importFrom officer docx_body_xml
process_par_settings <- function( rdoc ){
  all_nodes <- xml_find_all(docx_body_xml(rdoc), "//w:p/w:pPr[position()>1]")
  for(node_id in seq_along(all_nodes) ){
    pr <- all_nodes[[node_id]]
    par <- xml_parent(pr)
    pr1 <- xml_child(par, 1)
    if( xml_name(pr1) %in% "pPr" ){

      merge_pPr(pr, pr1, "w:pStyle")
      merge_pPr(pr, pr1, "w:jc")
      merge_pPr(pr, pr1, "w:spacing")
      merge_pPr(pr, pr1, "w:ind")
      merge_pPr(pr, pr1, "w:pBdr")
      merge_pPr(pr, pr1, "w:shd")

      xml_remove(pr)
    } else {
      xml_add_child(par, pr, .where = 1)
    }

  }
  rdoc
}


