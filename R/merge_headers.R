##' Unmerge cells that represent hierarchical headers and row labels.
##'
##' There is a pattern in headers where we have some number of levels
##' of merge, most commonly:
##'
##' \preformatted{
##'   | -------X------- |
##'   |  a  |  b  |  c  |
##' }
##'
##' which typically is intended to be interpreted as:
##'
##' \preformatted{
##'   | X:a | X:b | X:c |
##' }
##'
##' The same thing happen on rows, too.  I suspect that this is
##' generalisable to more than 2 columns.
##'
##' Within such header rows are either: vertically merged:
##'
##' \preformatted{
##'   | X | ===> | X |
##'   | | |
##' }
##'
##' or stacked
##'
##' \preformatted{
##'   | X | ===> | X:a |
##'   | a |
##' }
##'
##' or stacked and blank
##'
##' \preformatted{
##'   | X | or |   |  ===> | X |
##'   |   |    | X |
##' }
##'
##' I want to write something that handles this, that can be hooked up
##' to act as the "header" section on a view.  For now, let's apply it
##' to all cells within a range.
##' @title Unmerge headers and row labels
##' @param sheet A \code{\link{worksheet}} object
##' @param xr A cell range (a \code{cellranger::cell_limits} object)
##'   indicating the region to collapse.
##' @param horizontal Flag indicating if this a horizontal region
##'   representing headers (\code{TRUE}, the default) or a vertical
##'   region.  If horizontal we collapse away vertical spaces to
##'   create a single header row.  If vertical, we collapse away
##'   horizontal space to create a single row names column.
##' @param sep The separator to use between collapsed elements.  The
##'   default is a colon (\code{:}).  Because of the sheets this is
##'   designed to work with, you're never going to get syntactically
##'   valid names here, so feel free to use anything.  This can be
##'   multiple characters, a newline, whatever.  Likely you're going
##'   to have to process these names a bit later.
##' @return A character vector
##' @export
unmerge_header <- function(sheet, xr, horizontal=TRUE, sep=":") {
  ## TODO: check that a range fits within a sheet.  Should also be in
  ## cellranger?
  ##
  ## TODO: xr_to_idx also moves to cellranger?
  idx <- linen:::xr_to_idx(xr)
  i <- sheet$lookup[idx$r, idx$c, drop=FALSE]

  ## Now, we need to fill leftwards *or* downwards, depending on the
  ## `direction` argument.
  m <- array(vcapply(sheet$cells$value[i],
                     function(x) as.character(x %||% NA)),
             dim(i))

  ul <- xr$ul
  lr <- xr$lr

  for (el in sheet$merged) {
    ## TODO: This probably moves into linen:::loc_merge, but needs two things:
    ## 1. offset / anchor
    ## 2. horiz / vert only mode
    if (all(el$ul <= lr) && all(el$lr >= ul)) {
      if (horizontal) {
        if (el$lr[[2L]] > el$ul[[2L]]) {
          len <- el$lr[[2L]] - el$ul[[2L]] + 1L
          j <- cbind(el$ul[[1L]] - ul[[1L]] + 1L,
                     el$ul[[2L]] - ul[[2L]] + seq_len(len))
          m[j[-1L, , drop=FALSE]] <- m[j[1L, , drop=FALSE]]
        }
      } else {
        if (el$lr[[1L]] > el$ul[[1L]]) {
          len <- el$lr[[1L]] - el$ul[[1L]] + 1L
          j <- cbind(el$ul[[1L]] - ul[[1L]] + seq_len(len),
                     el$ul[[2L]] - ul[[2L]] + 1L)
          m[j[-1L, , drop=FALSE]] <- m[j[1L, , drop=FALSE]]
        }
      }
    }
  }

  apply(m, if (horizontal) 2 else 1,
        function(x) paste(na.omit(x), collapse=sep))
}
