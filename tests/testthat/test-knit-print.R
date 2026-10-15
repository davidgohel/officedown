rmd_lines <- c(
  "```{r}",
  "library(officedown)",
  "head(cars, 3)",
  "```"
)

test_that("data.frame is printed with knitr only", {
  rmd_file <- tempfile(fileext = ".Rmd")
  md_file <- tempfile(fileext = ".md")
  writeLines(rmd_lines, rmd_file)
  knitr::knit(rmd_file, output = md_file, quiet = TRUE, envir = new.env())
  md <- readLines(md_file)
  expect_false(any(grepl("length zero", md, fixed = TRUE)))
  expect_true(any(grepl("speed\\s+dist", md)))
})

test_that("data.frame is printed in non Office R Markdown outputs", {
  skip_if_not(rmarkdown::pandoc_available())
  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines(c("---", "output: md_document", "---", "", rmd_lines), rmd_file)
  md_file <- rmarkdown::render(rmd_file, quiet = TRUE, envir = new.env())
  expect_true(any(grepl("speed\\s+dist", readLines(md_file))))
})

test_that("data.frame does not corrupt other Word output formats", {
  skip_if_not(rmarkdown::pandoc_available())
  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines(c("---", "output: word_document", "---", "", rmd_lines), rmd_file)
  docx_file <- rmarkdown::render(rmd_file, quiet = TRUE, envir = new.env())
  dir_tmp <- tempfile()
  officer::unpack_folder(docx_file, dir_tmp)
  document <- readLines(file.path(dir_tmp, "word", "document.xml"), warn = FALSE)
  # officer temporary markup is only finalised by rdocx_document()
  expect_false(any(grepl("pstlname", document, fixed = TRUE)))
  expect_true(any(grepl("speed", document, fixed = TRUE)))
})

test_that("officer objects need rdocx_document in Word output", {
  skip_if_not(rmarkdown::pandoc_available())
  rmd_file <- tempfile(fileext = ".Rmd")
  writeLines(
    c(
      "---",
      "output: word_document",
      "---",
      "",
      "```{r}",
      "library(officedown)",
      "officer::fpar('a paragraph')",
      "```"
    ),
    rmd_file
  )
  expect_error(
    rmarkdown::render(rmd_file, quiet = TRUE, envir = new.env()),
    "officedown::rdocx_document"
  )
})
