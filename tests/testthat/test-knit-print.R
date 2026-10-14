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
