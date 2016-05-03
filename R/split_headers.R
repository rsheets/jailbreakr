##' Split headers from a sheet or a view.  This function is designed
##' to work well with \code{\link{merge_headers}}, as some of the
##' headers that will be detected here are going to be multi-row and
##' therefore need collapsing.
##'
##' Like \code{\link{split_metadata}}, we need to employ some
##' heuristics to get this right.  These will probably evolve as we
##' throw them at more spreadsheets.
##'
##' This seems pretty hard to generalise, but at the same time it's
##' easy enough to tell when \emph{looking} at a spreadsheet which
##' bits are headers, but it seems hard to teach the package what
##' these rules are.
##'
##' Things that are going to be common include varying the background
##' colour; to detect that we're going to need some colour distance
##' metrics (plain colour shifts aren't enough because multi-colour
##' headings are going to be common).  Deciding what the relevant
##' colour differences is tricky because that will depend on the sheet
##' so I should look into perhaps something about edge detection here?
##'
##' Another common way of marking out header rows will be boldness.
##' That often will carry through to the row labels too though, but
##' we'd be OK looking for rows where essentially everything is bold
##' vs ones where a few things are bold.
##'
##' If we're lucky enough to have split frames \emph{and} we have a
##' full sheet (not a view) then we can probably treat that as the
##' point where the headers start.
##'
##' Otherwise we can filter based on things like 
##' 
##' @title Split headers from a sheet or a view
##' @param x A worksheet or a worksheet view
##' @param data_allow_merged If \code{TRUE}, horizontal merged cells
##'   are allowed in data regions, otherwise a region is considered
##'   not to be a data region if it has any merged cells that extend
##'   horizontally.  Vertical merged cells, so long as they do not
##'   span horizontally, are allowed.
##' @return For \code{split_headers} and \code{split_headers_apply},
##'   a worksheet view; in this view the \code{data$headers} element
##'   will be a worksheet view of the headers.  For
##'   \code{split_headers_find}, a single integer representing the
##'   number of rows of headers found (with zero indicating no
##'   headers).
##' @export
split_headers <- function(x, data_allow_merged=FALSE) {
  stop("Not implemented")
}

##' @export
##' @rdname split_headers
split_headers_find <- function(x, use_frames=TRUE,

                               data_allow_merged=FALSE) {
  
}

##' @export
##' @rdname split_headers
##' @param n Single integer indicating how many rows are considered to
##'   be headers.
split_headers_apply <- function(x, n) {
  dim <- x$dim
  nr <- dim[[1L]]
  nc <- dim[[2L]]
  nr <- x$dim[[2L]]

  xr <- x$xr
  if (is.null(xr)) {
    xr <- cellranger::cell_limits(c(1, 1), c(nr - n, nc))
  } else {
    xr$ul[[1]] <- xr$ul[[1]] + n
  }

  header <- unmerge_headers(x, n)
  ## may make headers first class elements of worksheet_views, not sure.
  linen::worksheet_view(x$sheet, xr, header=header, data=x$data)
}
