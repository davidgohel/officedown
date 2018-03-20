#' @import officer xml2
#' @importFrom htmltools htmlEscape
process_images <- function( rdoc ){
  rel <- rdoc$doc_obj$relationship()
  blips <- xml_find_all(rdoc$doc_obj$get(), "//a:blip[@r:embed]")
  invalid_blips <- blips[!grepl( "^rId[0-9]+$", xml_attr(blips, "embed") )]
  image_paths <- unique(xml_attr(invalid_blips, "embed") )

  for(i in seq_along(image_paths) ){

    rid <- sprintf("rId%.0f", rel$get_next_id() )

    img_dir <- file.path(rdoc$doc_obj$package_dirname(), "word", "media")
    dir.create(img_dir, recursive = TRUE, showWarnings = FALSE)

    new_img_path <- basename(tempfile(fileext = gsub("(.*)(\\.[0-9a-zA-Z]+$)", "\\2", image_paths[i])))
    new_img_path <- file.path(img_dir, new_img_path)
    file.copy(image_paths[i], to = new_img_path)

    rel$add_img(new_img_path, root_target = "media")
    which_match_path <- grepl( image_paths[i], xml_attr(invalid_blips, "embed"), fixed = TRUE )
    xml_attr(invalid_blips[which_match_path], "r:embed") <- rep(rid, sum(which_match_path))
  }
  rdoc
}

process_links <- function( rdoc ){
  rel <- rdoc$doc_obj$relationship()
  hl_nodes <- xml_find_all(rdoc$doc_obj$get(), "//w:hyperlink[@r:id]")
  which_to_add <- hl_nodes[!grepl( "^rId[0-9]+$", xml_attr(hl_nodes, "id") )]
  hl_ref <- unique(xml_attr(which_to_add, "id"))
  for(i in seq_along(hl_ref) ){
    rid <- sprintf("rId%.0f", rel$get_next_id() )

    rel$add(
      id = rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target = htmlEscape(hl_ref[i]), target_mode = "External" )

    which_match_id <- grepl( hl_ref[i], xml_attr(which_to_add, "id"), fixed = TRUE )
    xml_attr(which_to_add[which_match_id], "r:id") <- rep(rid, sum(which_match_id))
  }
  rdoc
}

process_chunk_style <- function( rdoc ){
  styles_table <- styles_info(rdoc)

  all_nodes <- xml_find_all(rdoc$doc_obj$get(), "//w:rStyle[@w:chunk_style]")

  ref_styles <- unique(xml_attr(all_nodes, "chunk_style"))
  for(i in seq_along(ref_styles) ){

    style_name <- ref_styles[i]
    which_match_id <- grepl( style_name, xml_attr(all_nodes, "chunk_style"), fixed = TRUE )

    if( !style_name %in% styles_table$style_name ){
      warning("could not match any style named ", shQuote(style_name), call. = FALSE)
      xml_remove(all_nodes[which_match_id])
      break
    }

    style_id <- styles_table$style_id[styles_table$style_name %in% style_name]

    xml_attr(all_nodes[which_match_id], "w:val") <- rep(style_id, sum(which_match_id))
    xml_attr(all_nodes[which_match_id], "w:chunk_style") <- NULL

  }
  rdoc
}
correct_styles <- function( rdoc, fig_cap, fig, title, author, date ){
  styles_table <- styles_info(rdoc)
  pars <- c(ImageCaption = fig_cap, CaptionedFigure = fig,
            Title = title, Author = author, Date = date)
  for( what in names(pars) ){
    match_id <- which( styles_table$style_type %in% "paragraph" & styles_table$style_name %in% pars[what] )
    if( length(match_id) > 0 ){
      styleid <- styles_table$style_id[match_id]
      all_nodes <- xml_find_all(rdoc$doc_obj$get(), sprintf("//w:pStyle[@w:val='%s']", what))
      xml_attr(all_nodes, "w:val") <- rep(styleid, length(all_nodes) )
    } else {
      warning("could not find any style named ", shQuote(pars[what]), ".", call. = FALSE)
    }
  }
  rdoc
}

process_parinst <- function( rdoc ){

  all_nodes <- xml_find_all(rdoc$doc_obj$get(), "//w:pPrworded")
  for(node_id in seq_along(all_nodes) ){
    par <- xml_parent(all_nodes[[node_id]])

    jc <- xml_child(all_nodes[[node_id]], "w:jc")
    jc_ref <- xml_child(par, "w:jc")
    if( !inherits(jc, "xml_missing") ){
      if( !inherits(jc_ref, "xml_missing") )
        jc_ref <- xml_replace(jc_ref, jc)
      else xml_add_child(xml_child(par, "w:pPr"), jc)
    }


    spacing <- xml_child(all_nodes[[node_id]], "w:spacing")
    if( !inherits(spacing, "xml_missing") ){
      spacing_ref <- xml_child(par, "w:pPr/w:spacing")
      if( !inherits(spacing_ref, "xml_missing") )
        spacing_ref <- xml_replace(spacing_ref, spacing)
      else xml_add_child(xml_child(par, "w:pPr"), spacing)
    }

    indent <- xml_child(all_nodes[[node_id]], "w:ind")
    if( !inherits(indent, "xml_missing") ){
      indent_ref <- xml_child(par, "w:pPr/w:ind")
      if( !inherits(indent_ref, "xml_missing") )
        indent_ref <- xml_replace(indent_ref, indent)
      else xml_add_child(xml_child(par, "w:pPr"), indent)
    }

    bb <- xml_child(all_nodes[[node_id]], "w:pBdr")
    if( !inherits(bb, "xml_missing") ){
      bb_ref <- xml_child(par, "w:pPr/w:pBdr")
      if( !inherits(bb_ref, "xml_missing") )
        bb_ref <- xml_replace(bb_ref, bb)
      else xml_add_child(xml_child(par, "w:pPr"), bb)
    }

    shd <- xml_child(all_nodes[[node_id]], "w:shd")
    if( !inherits(shd, "xml_missing") ){
      shd_ref <- xml_child(par, "w:pPr/w:shd")
      if( !inherits(shd_ref, "xml_missing") )
        shd_ref <- xml_replace(shd_ref, shd)
      else xml_add_child(xml_child(par, "w:pPr"), shd)
    }

    xml_remove(all_nodes[[node_id]])
  }
  rdoc
}

process_sections <- function( rdoc ){

  all_nodes <- xml_find_all(rdoc$doc_obj$get(), "//w:sectPr[w:worded]")
  main_sect <- xml_find_first(rdoc$doc_obj$get(), "w:body/w:sectPr")

  for(node_id in seq_along(all_nodes) ){
    current_node <- as_xml_document(all_nodes[[node_id]])
    new_node <- as_xml_document(main_sect)

    # correct type ---
    type <- xml_child(current_node, "w:type")
    type_ref <- xml_child(new_node, "w:type")
    if( !inherits(type, "xml_missing") ){
      if( !inherits(type_ref, "xml_missing") )
        type_ref <- xml_replace(type_ref, type)
      else xml_add_child(new_node, type)
    }

    # correct cols ---
    cols <- xml_child(current_node, "w:cols")
    cols_ref <- xml_child(new_node, "w:cols")
    if( !inherits(cols, "xml_missing") ){
      if( !inherits(cols_ref, "xml_missing") )
        cols_ref <- xml_replace(cols_ref, cols)
      else xml_add_child(new_node, cols)
    }

    # correct pgSz ---
    pgSz <- xml_child(current_node, "w:pgSz")
    pgSz_ref <- xml_child(new_node, "w:pgSz")
    if( !inherits(pgSz, "xml_missing") ){

      if( !inherits(pgSz_ref, "xml_missing") ){
        xml_attr(pgSz_ref, "w:orient") <- xml_attr(pgSz, "orient")

        wref <- as.integer( xml_attr(pgSz_ref, "w") )
        href <- as.integer( xml_attr(pgSz_ref, "h") )

        if( xml_attr(pgSz, "orient") %in% "portrait" ){
          h <- ifelse( wref < href, href, wref )
          w <- ifelse( wref < href, wref, href )
        } else if( xml_attr(pgSz, "orient") %in% "landscape" ){
          w <- ifelse( wref < href, href, wref )
          h <- ifelse( wref < href, wref, href )
        }
        xml_attr(pgSz_ref, "w:w") <- w
        xml_attr(pgSz_ref, "w:h") <- h
      } else {
        xml_add_child(new_node, pgSz)
      }
    }

    node <- xml_replace(all_nodes[[node_id]], new_node)
  }
  rdoc
}

process_embedded_docx <- function( rdoc ){
  rel <- rdoc$doc_obj$relationship()
  hl_nodes <- xml_find_all(rdoc$doc_obj$get(), "//w:altChunk[@r:id]")
  which_to_add <- hl_nodes[!grepl( "^rId[0-9]+$", xml_attr(hl_nodes, "id") )]
  hl_ref <- unique(xml_attr(which_to_add, "id"))

  for(i in seq_along(hl_ref) ){
    which_match_id <- grepl( hl_ref[i], xml_attr(which_to_add, "id"), fixed = TRUE )

    if( !file.exists(hl_ref[i]) ){
      for( n in seq_along(which_to_add[which_match_id] )) xml_remove(which_to_add[which_match_id][[n]] )
      break
    }

    rid <- sprintf("rId%.0f", rel$get_next_id() )
    new_docx_file <- basename(tempfile(fileext = ".docx"))

    file.copy(hl_ref[i], to = file.path(rdoc$package_dir, new_docx_file))
    rel$add(
      id = rid, type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
      target = file.path("../", new_docx_file) )

    override <- paste0("/", new_docx_file)
    names(override) <- "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    rdoc$content_type$add_override( override )
    xml_attr(which_to_add[which_match_id], "r:id") <- rep(rid, sum(which_match_id))

  }
  rdoc
}

# macro ----
LIST_CHUNK_MACRO <- list(
  MACRO_PAGEBREAK = chunk_page_break,
  MACRO_COLUMNBREAK = chunk_column_break
  )
LIST_CHUNK_MACRO_PAR <- list(
  MACRO_PAR_SETTING = "macro_paragraph_format"
  )

LIST_BLOCK_MACRO <- list(
  MACRO_TOC = "block_toc",
  MACRO_POUR_DOCX = "block_docx_pour",
  MACRO_LANDSCAPE_START = "macro_section_continuous",
  MACRO_LANDSCAPE_STOP = "macro_section_landscape",
  MACRO_MULTICOL_START = "macro_section_continuous",
  MACRO_MULTICOL_STOP = "macro_section_columns"
)

chunk_macro <- function(txt, type = "docx"){
  for( i in names(LIST_CHUNK_MACRO) ){
    txt <- gsub(paste0("<!---", i, "--->"), format(LIST_CHUNK_MACRO[[i]](), type = type), txt)
  }
  txt
}
#' @import stringr
chunk_macro_par <- function(txt, type = "docx"){
  for( i in names(LIST_CHUNK_MACRO_PAR) ){
    regex <- paste0("<!---", i, "[ ]*(\\{.*\\}){0,1}[ ]*--->")
    reg_ <- regexpr(pattern = regex, text = txt)
    if( any( macro_ <- reg_ > -1 ) ){
      newpars <- substr(txt[macro_], reg_[macro_], stop = reg_[macro_] + attr(reg_, "match.length")[macro_])
      newpars <- gsub(paste0("<!---", i, "[ ]*\\{"), "", newpars)
      newpars <- gsub(paste0("\\}--->"), "", newpars)
      newpars <- sapply( newpars, function(x, fname, type){
        x <- eval(parse( text = paste0( fname, "(", x, ")") ) )
        format(x, type = type)
      },
      fname = LIST_CHUNK_MACRO_PAR[[i]], type = type )
      txt[macro_] <- str_replace(txt[macro_], pattern = regex, replacement = newpars)
    }
  }
  txt
}


block_macro <- function(txt, type = "docx"){
  for( i in names(LIST_BLOCK_MACRO) ){
    regex <- paste0("^([ ]*<!---)(", i, ")([ ]*)(\\{.*\\}){0,1}(--->[ ]*)$")
    if( any( macro_ <- grepl(regex, txt) ) ){
      txt[macro_] <- sapply( gsub(regex, "\\4", txt)[macro_], function(x, fname, type){
        x <- gsub("(^\\{|\\}$)", "", x)
        x <- eval(parse( text = paste0( fname, "(", x, ")") ) )
        format(x, type = type)
      },
      fname = LIST_BLOCK_MACRO[[i]], type = type )
    }
  }
  txt
}

# rdocx_document ----
#' @export
rdocx_document <- function(fig_cap = "Image Caption", fig = "Captioned Figure",
                           title = "Title", author = "Author", date = "Date", ...) {

  output_formats <- rmarkdown::word_document(...)

  output_formats$pre_processor =  function(metadata, input_file, runtime, knit_meta, files_dir, output_dir){
    md <- readLines(input_file)
    md <- chunk_macro(md)
    md <- chunk_macro_par(md)
    md <- block_macro(md)
    writeLines(md, input_file)
  }

  output_formats$post_processor <- function(metadata, input_file, output_file, clean, verbose) {
    x <- officer::read_docx(output_file)
    x <- process_images(x)
    x <- process_links(x)
    x <- process_embedded_docx(x)
    x <- process_chunk_style(x)
    x <- process_sections(x)
    x <- process_parinst(x)
    x <- correct_styles(x, fig_cap = fig_cap, fig = fig, title = title, author = author, date = date)

    print(x, target = output_file)
    output_file
  }
  output_formats
}


