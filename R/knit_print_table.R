sanitize_tab.lp <- function(tab.lp){
  if(is.null(tab.lp)){
    tab.lp <- "tab"
  } else {
    tab.lp <- gsub("[:]{1}$", "", tab.lp)
  }
  tab.lp
}


default_style <- function(type, si){
  si[si$style_type %in% type & si$is_default ,"style_name"]
}
style_id <- function(x, type, si){
  si[
    si$style_type %in% type &
      si$style_name %in% x ,
    "style_id"]
}
validate_style <- function(x, type, si){
  validated_style <- si[si$style_type %in% type & si$style_name %in% x, "style_name"]
  if(length(validated_style) != 1 ){
    validated_style <- default_style(type, si)
    msg <- paste0("could not find ", type, " style ", shQuote(x),
                  ". Switching to default one named ", shQuote(validated_style), ".")
    warning(msg, call. = FALSE)
  }
  validated_style
}

get_table_design_opt <- function(x, default = FALSE){
  x <- opts_current$get(x)
  if(is.null(x)) x <- default
  x
}


# knit_print.data.frame -----

#' @importFrom officer block_table prop_table table_layout table_width table_colwidths table_conditional_formatting
#'  opts_current_table block_caption styles_info
#' @importFrom knitr knit_print asis_output opts_current
knit_print.data.frame <- function(x, ...) {

  tab_props <- opts_current_table()
  if( grepl( "docx", opts_knit$get("rmarkdown.pandoc.to") ) ){

    opts_knit$get("rmarkdown.pandoc.to")
    pt <- prop_table(
      style = tab_props$style, layout = tab_props$table_layout,
      width = tab_props$table_width,
      tcf = table_conditional_formatting(
        first_row = tab_props$first_row,
        first_column = tab_props$first_column,
        last_row = tab_props$last_row,
        last_column = tab_props$last_column,
        no_hband = tab_props$no_hband,
        no_vband = tab_props$no_vband))

    bt <- block_table(x,
                      header = get_table_design_opt("header", default = TRUE),
                      properties = pt
                      )
    bc <- block_caption(label = tab_props$cap, style = tab_props$cap.style,
                  autonum = run_autonum(
                    seq_id = gsub(":$", "", tab_props$tab.lp),
                    pre_label = tab_props$cap.pre,
                    post_label = tab_props$cap.sep
                  ))

    cap_str <- to_wml(bc, knitting = TRUE)
    res <- paste(cap_str, "```{=openxml}",
                 to_wml(bt, base_document = get_reference_rdocx()),
                 "```\n\n",
                 sep = "\n")
    asis_output(res)
  } else if( grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to") ) ){

    layout <- knitr::opts_current$get("layout")
    master <- knitr::opts_current$get("master")
    doc <- get_reference_pptx()

    if(is.null( ph <- knitr::opts_current$get("ph") )){
      ph <- officer::ph_location_type(type = "body")
    }
    if(!inherits(ph, "location_str")){
      stop("ph should be a placeholder location; ",
           "see officer::placeholder location for an example.",
           call. = FALSE)
    }

    location <- get_content_ph(ph, layout, master, doc)

    pt <- prop_table(style = doc$table_styles$def[1],
      tcf = table_conditional_formatting(
        first_row = tab_props$first_row,
        first_column = tab_props$first_column,
        last_row = tab_props$last_row,
        last_column = tab_props$last_column,
        no_hband = tab_props$no_hband,
        no_vband = tab_props$no_vband))

    bt <- block_table(x,
                      header = get_table_design_opt("header", default = TRUE),
                      properties = pt)
    res <- paste("```{=openxml}",
                 officer::to_pml(bt, left = location$left, top = location$top,
                                 width = location$width, height = location$height,
                                 label = location$ph_label, ph = location$ph,
                                 rot = location$rotation, bg = location$bg),
                 "```\n\n",
                 sep = "\n")
    asis_output(res)
  } else knit_print( asis_output("") )
}

