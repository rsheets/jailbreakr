## Needed for reference comparisons
dir.create("reference", FALSE)

## This is needed because R installs without including the tests by
## default.
get_readxl <- function(path="readxl") {
  if (file.exists(path)) {
    return(path)
  }
  url <- "https://cran.rstudio.com/src/contrib/readxl_0.1.0.tar.gz"
  dest <- tempfile()
  tryCatch(download.file(url, dest),
           error=function(e) skip(e$message))
  on.exit(file.remove(dest))
  untar(dest, path)
  path
}
