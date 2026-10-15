# Advanced R Markdown PowerPoint Format

Format for converting from R Markdown to an MS PowerPoint document.

The function will allow you to specify the destination of your chunks in
the output PowerPoint file. In this case, you must specify the `layout`
and `master` for the layout you want to use, as well as the `ph`
argument, which will allow you to specify the placeholder to be
generated to place the result. Use the officer package to help you
choose the identfiers to use.

This function also support Vector graphics output in an editable format
(using package `rvg`). Wrap you R plot commands with function `dml` to
use this graphic capability.

## Usage

``` r
rpptx_document(
  base_format = "rmarkdown::powerpoint_presentation",
  layout = "Title and Content",
  master = "Office Theme",
  tcf = list(),
  ...
)
```

## Arguments

- base_format:

  a scalar character, format to be used as a base document for
  officedown. default to
  [powerpoint_presentation](https://pkgs.rstudio.com/rmarkdown/reference/powerpoint_presentation.html)
  but can also be powerpoint_presentation2 from bookdown

- layout:

  default slide layout name to use

- master:

  default master layout name where `layout` is located

- tcf:

  default conditional formatting settings defined by
  [`officer::table_conditional_formatting()`](https://davidgohel.github.io/officer/reference/table_conditional_formatting.html)

- ...:

  arguments used by
  [powerpoint_presentation](https://pkgs.rstudio.com/rmarkdown/reference/powerpoint_presentation.html)

## Value

R Markdown output format to pass to
[render](https://pkgs.rstudio.com/rmarkdown/reference/render.html)

## Examples

``` r
library(rmarkdown)
run_ok <- pandoc_available() && pandoc_version() > numeric_version("2.4")

if(run_ok){
  example <- system.file(package = "officedown",
    "examples/minimal_powerpoint.Rmd")
  rmd_file <- tempfile(fileext = ".Rmd")
  file.copy(example, to = rmd_file)

  pptx_file_1 <- tempfile(fileext = ".pptx")
  # \donttest{
  render(rmd_file, output_file = pptx_file_1)
  # }
}
#> 
#> 
#> processing file: file1a332952fcc6.Rmd
#> 1/14                  
#> 2/14 [setup]          
#> 3/14                  
#> 4/14 [unnamed-chunk-1]
#> 5/14                  
#> 6/14 [unnamed-chunk-2]
#> 7/14                  
#> 8/14 [plot1]          
#> 9/14                  
#> 10/14 [unnamed-chunk-3]
#> 11/14                  
#> 12/14 [unnamed-chunk-4]
#> 13/14                  
#> 14/14 [unnamed-chunk-5]
#> output file: file1a332952fcc6.knit.md
#> /opt/hostedtoolcache/pandoc/3.8.3/x64/pandoc +RTS -K512m -RTS file1a332952fcc6.knit.md --to pptx --from markdown+autolink_bare_uris+tex_math_single_backslash --output /tmp/RtmpE4B43W/file1a334521eb68.pptx --reference-doc /tmp/RtmpE4B43W/file1a33fc82ac8.pptx --lua-filter /home/runner/work/_temp/Library/rmarkdown/rmarkdown/lua/extract-data-uri.lua 
#> 
#> Output created: /tmp/RtmpE4B43W/file1a334521eb68.pptx
```
