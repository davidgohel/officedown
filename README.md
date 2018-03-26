
<!-- README.md is generated from README.Rmd. Please edit that file -->

## Usage

use RStudio Menu to create a document from `worded` template.

![](tools/rstudio_new_rmd.gif)

It will create an R markdown document, parameter `output` is to be set
to `worded::rdocx_document`. Also package `worded` need to be loaded.

<pre>
---
date: "2018-03-26"
author: "David Gohel"
title: "Document title"
output: 
  worded::rdocx_document
---

&#96;&#96;&#96;{r setup, include=FALSE}
library(worded)
&#96;&#96;&#96;

...

</pre>

## Available functions

### Chunks

Chunks are to be used in a paragraph in an R markdown
document.

#### Break page

<pre>This will add a break page: <!--html_preserve--><span style="color:red;">&lt;!---CHUNK_PAGEBREAK---&gt;</span><!--/html_preserve-->
It is also possible to use inline R expression <code>&#96;<!--html_preserve--><span style="color:red;">r chunk_page_break()</span><!--/html_preserve-->&#96;</code>.
</pre>

#### Styled text

<pre>This will add a styled text: <!--html_preserve--><span style="color:red;">&lt;!---CHUNK_TEXT{str: 'text', color: 'orange'}---&gt;</span><!--/html_preserve-->
It is also possible to use inline R expression <code>&#96;<!--html_preserve--><span style="color:red;">r chunk_styled_text('text', color = 'orange')</span><!--/html_preserve-->&#96;</code>.
</pre>

#### Associate text with a style name

<pre>This will add a styled text: <!--html_preserve--><span style="color:red;">&lt;!---CHUNK_TEXT_STYLE{str: 'text', style: 'redstrong'}---&gt;</span><!--/html_preserve-->
It is also possible to use inline R expression <code>&#96;<!--html_preserve--><span style="color:red;">r chunk_text_stylenamed('text', style = 'redstrong')</span><!--/html_preserve-->&#96;</code>.
</pre>

#### Break column

Break column has to be used with a section with multiple
columns.

<pre>This will add a break column: <!--html_preserve--><span style="color:red;">&lt;!---CHUNK_COLUMNBREAK---&gt;</span><!--/html_preserve-->
It is also possible to use inline R expression <code>&#96;<!--html_preserve--><span style="color:red;">r chunk_column_break()</span><!--/html_preserve-->&#96;</code>.
</pre>

### Blocks

## Installation

You can install worded from github with:

``` r
# install.packages("devtools")
devtools::install_github("davidgohel/worded")
```
