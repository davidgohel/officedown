render_rmd <- function(rmd_file, output_file) {
  unlink(output_file, force = TRUE)
  sucess <- FALSE
  tryCatch(
    {
      rmarkdown::render(
        rmd_file,
        output_file = output_file,
        envir = new.env(),
        quiet = TRUE
      )
      sucess <- TRUE
    },
    warning = function(e) {},
    error = function(e) {}
  )
  sucess
}
get_docx_xml <- function(x) {
  redoc <- read_docx(x)
  xml2::xml_child(docx_body_xml(redoc))
}

# Build a reference docx whose default section carries a header and a
# footer, so it can be used as `reference_docx` in tests (#127).
ref_docx_with_hf <- function() {
  sect <- officer::prop_section(
    type = "continuous",
    header_default = officer::block_list(officer::fpar("REF_HEADER_TEXT")),
    footer_default = officer::block_list(officer::fpar("REF_FOOTER_TEXT"))
  )
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "placeholder", style = "Normal")
  doc <- officer::body_set_default_section(doc, sect)
  out <- tempfile(fileext = ".docx")
  print(doc, target = out)
  out
}
