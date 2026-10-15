# Force Block Printing while Knitting

When used in a loop, calls to blocks do not generate output because
`knit_print` method is not called. Use the function to force printing.
Also you should tell the chunk to use results 'as-is' (by adding
`results='asis'` to your chunk header).

## Usage

``` r
knit_print_block(x, ...)
```

## Arguments

- x:

  a block object, result of a block function from officer package

- ...:

  unused arguments

## Value

None. the function only print XML code.

## See also

Other functions that force printing:
[`knit_print_run()`](https://davidgohel.github.io/officedown/dev/reference/knit_print_run.md)

## Examples

``` r
library(rmarkdown)
rmd_file_src <- system.file(
  package = "officedown", "examples", "word_loop.Rmd")
rmd_file_des <- tempfile(fileext = ".Rmd")
if(pandoc_available()){

  file.copy(rmd_file_src, to = rmd_file_des)
  docx_file_1 <- tempfile(fileext = ".docx")
  render(rmd_file_des, output_file = docx_file_1, quiet = TRUE)

  if(file.exists(docx_file_1)){
    message("file ", docx_file_1, " has been written.")
  }
}
#> file /tmp/RtmpXy0Iqe/file1a53382f1428.docx has been written.
```
