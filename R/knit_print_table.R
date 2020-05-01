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

#' @importFrom officer styles_info
opts_current_table <- function(){
  tab.cap.style <- get_table_design_opt("tab.cap.style", default = "Normal")
  tab.cap.pre <- get_table_design_opt("tab.cap.pre", default = "Table ")
  tab.cap.sep <- get_table_design_opt("tab.cap.sep", default = ":")
  tab.cap <- opts_current$get("tab.cap")
  tab.id <- opts_current$get("tab.id")
  tab.style <- opts_current$get("tab.style")
  tab.layout <- get_table_design_opt("tab.layout", default = "autofit")
  tab.width <- get_table_design_opt("tab.width", default = 1)

  first_row <- get_table_design_opt("first_row")
  first_column <- get_table_design_opt("first_column")
  last_row <- get_table_design_opt("last_row")
  last_column <- get_table_design_opt("last_column")
  no_hband <- get_table_design_opt("no_hband")
  no_vband <- get_table_design_opt("no_vband")

  doc <- get_reference_rdocx()
  si <- styles_info(doc)

  if(is.null(tab.cap.style)){
    tab.cap.style <- default_style("paragraph", si)
  } else {
    tab.cap.style <- validate_style(x = tab.cap.style, type = "paragraph", si = si)
  }
  tab.cap.style_id <- style_id(tab.cap.style, type = "paragraph", si)

  if(is.null(tab.cap.pre)){
    tab.cap.pre <- "table "
  }
  if(is.null(tab.cap.sep)){
    tab.cap.sep <- ": "
  }
  if(is.null(tab.layout)){
    tab.layout <- "autofit"
  }
  if(is.null(tab.width)){
    tab.width <- 1
  }

  if(is.null(tab.style)){
    tab.style <- default_style("table", si)
  } else {
    tab.style <- validate_style(x = tab.style, type = "table", si = si)
  }


  list(cap.style = tab.cap.style, cap.style_id = tab.cap.style_id,
       cap.pre = tab.cap.pre, cap.sep = tab.cap.sep,
       id = tab.id, cap = tab.cap,
       style = tab.style, seq_id = "tab",
       table_layout = table_layout(type = tab.layout),
       table_width = table_width(width = tab.width, unit = "pct"),
       first_row = first_row,
       first_column = first_column,
       last_row = last_row,
       last_column = last_column,
       no_hband = no_hband,
       no_vband = no_vband
       )

}

# knit_print.data.frame -----

#' @importFrom officer block_table prop_table table_layout table_width table_colwidths table_conditional_formatting
#' @importFrom knitr knit_print asis_output opts_current
knit_print.data.frame <- function(x, ...) {

  if( grepl( "docx", opts_knit$get("rmarkdown.pandoc.to") ) ){
    tab_props <- opts_current_table()
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

    cap_str <- do.call(pandoc_wml_caption, tab_props)
    res <- paste(cap_str, "```{=openxml}",
                 to_wml(bt, base_document = get_reference_rdocx()),
                 "```\n\n",
                 sep = "\n")
    asis_output(res)
  } else if( grepl( "pptx", opts_knit$get("rmarkdown.pandoc.to") ) ){
    if(is.null( layout <- knitr::opts_current$get("layout") )){
      layout <- officer::ph_location_type()
    }
    location <- get_content_layout(layout)

    tab_props <- opts_current_table()
    bt <- block_table(x)
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

