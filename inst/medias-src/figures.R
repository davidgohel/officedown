library(doconv)
library(rmarkdown)
library(rsvg)

# bookdown patchwork -----
dir_src <- "inst/medias-src/figures"
docx_filename <- file.path("bookdown", "_book", "bookdown.docx")
docx_filename_dest <- file.path(dir_src, "bookdown.docx")


bookdown_loc <- system.file(package = "officedown", "examples/bookdown")
file.copy(from = bookdown_loc, to = getwd(), overwrite = TRUE, recursive = TRUE)
render_site(input = "bookdown", encoding = 'UTF-8', quiet = TRUE)
file.copy(docx_filename, to = docx_filename_dest)
unlink("bookdown", recursive = TRUE, force = TRUE)
to_miniature(filename = docx_filename_dest, use_docx2pdf = TRUE,
             row = c(1, 1, 2, 2, 3, 0, 4, 4, 5, 5, 6), width = 850,
             fileout = file.path(dir_src, "README-bookdown.png"))

unlink(list.files(dir_src, pattern = "^bookdown", full.names = TRUE), force = TRUE)

# logos -----

rsvg_png(svg = "inst/medias-src/logo-src.svg",
         file = "inst/medias-src/figures/logo.png",
         width = 200, height = 231)

# compression -----
minimage::compress_images(input = "inst/medias-src/figures", output = "man/figures")

