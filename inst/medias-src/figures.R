library(magrittr)
library(purrr)
library(magick)
library(pdftools)
library(imgpress)
library(rmarkdown)
library(rsvg)

# bookdown patchwork -----
# i/o
dir_src <- "inst/medias-src/figures"
docx_filename <- file.path("bookdown", "_book", "bookdown.docx")
docx_filename_dest <- file.path(dir_src, "bookdown.docx")
pdf_filename_dest <- file.path(dir_src, "bookdown.pdf")


bookdown_loc <- system.file(package = "officedown",
                            "examples/bookdown")

file.copy(from = bookdown_loc, to = getwd(),
          overwrite = TRUE, recursive = TRUE)

render_site(input = "bookdown",
            encoding = 'UTF-8',
            quiet = TRUE)
file.copy(docx_filename, to = docx_filename_dest)
unlink("bookdown", recursive = TRUE, force = TRUE)

browseURL(docx_filename_dest)
readline(prompt = "type any key when document is updated and saved as pdf > ")

unlink(docx_filename_dest, force = TRUE)

pdftools::pdf_convert(pdf_filename_dest,
                      filenames = file.path(dir_src, "bookdown_%02d.%s"),
                      format = "png")
unlink(pdf_filename_dest, force = TRUE)

img_list <- list.files(dir_src, pattern = "^bookdown_", full.names = TRUE) %>%
  map(image_read) %>% map(image_border, color = "#CCC", geometry = "2x2")

img_stack <- list()
img_stack[[1]] <- image_append( c(img_list[[1]], img_list[[2]]) ) %>% image_extent(geometry = "1674x")
img_stack[[2]] <- image_append( c(img_list[[3]], img_list[[4]]) )  %>% image_extent(geometry = "1674x")
img_stack[[3]] <- img_list[[5]] %>% image_extent(geometry = "1674x")
img_stack[[4]] <- image_append( c(img_list[[7]], img_list[[8]]) )  %>% image_extent(geometry = "1674x")
img_stack[[5]] <- img_list[[9]] %>% image_extent(geometry = "1674x")

do.call(c, img_stack) %>%
  image_append(stack = TRUE) %>%
  image_write(path = file.path(dir_src, "README-bookdown.png"))

list.files(dir_src, pattern = "^bookdown_", full.names = TRUE) %>%
  unlink(force = TRUE)

# logos -----

rsvg_png("inst/medias-src/logo-src.svg", file = "inst/medias-src/figures/officedown.png", width = 200, height = 231)
rsvg_png("inst/medias-src/logo-src-inv.svg", file = "inst/medias-src/figures/logo.png", width = 200, height = 231)

# compression -----
imgpress::img_compress(dir_input = "inst/medias-src/figures",
                       dir_output = "man/figures")

