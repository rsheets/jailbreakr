## This is surprisingly hard to do in the general case because we
## need to find all islands that satisfy the general property; they
## might not be aligned so we can't use the general approach of
## looking for blank rows or columns (no matter how much that
## *seems* like it would be enough) because if the sub-tables do not
## align then this will miss sub tables.

## Define a subtable as a possibly ragged block of cells surrounded by
## an edge of the table, or by a set of blank cells.  I believe that
## the "flood fill" algorithm here is a reasonable choice and using a
## 4-direction queue will be fairly efficient.

##' Classify a table into sub-regions.  We're looking for a (possibly
##' ragged) block of cells surrounded by a set of blank cells or the
##' edge of the sheet.  This does not detect regions that are multiple
##' square regions offset from each other, but that could possibly be
##' done afterwards.
##'
##' This function works by applying the "flood fill" algorithm to
##' non-blank cells in the worksheet and then squaring off the result.
##'
##' The \code{classify_sheet} function does the actual classification,
##' and \code{split_sheet} applies \code{worksheet_view} to these to
##' produce something that can be used (approximately) as if it was a
##' separate sheet.
##' @title Classify and split sheet
##' @param sheet A linen \code{worksheet} object, possibly from an
##'   Excel or googlesheets spreadsheet.
##' @param as Character, indicating what to return - \code{"limits"}:
##'   a list of limits (the default), \code{"groups"} a matrix of the
##'   same dimensions as the worksheet indicating what group each cell
##'   is in, or \code{"both"}: a list with elements \code{"limits"}
##'   and \code{"groups"}.
##' @export
classify_sheet <- function(sheet, as="limits") {
  if (!inherits(sheet, "worksheet")) {
    stop("sheet must be a 'worksheet' object")
  }
  as <- match.arg(as, c("limits", "groups", "both"))

  i <- abs(sheet$lookup2)
  i <- !is.na(i) & !sheet$cells$is_blank[c(i)]
  ## writeLines(apply(ifelse(i, "*", " "), 1, paste, collapse=""))

  ## Pad the array with FALSE to avoid a lot of conditional switches.
  nc <- ncol(i)
  nr <- nrow(i)
  x <- matrix(FALSE, nr + 2L, nc + 2L)
  x[seq_len(nr) + 1L, seq_len(nc) + 1L] <- i
  valid <- array(TRUE, dim(x))
  grp <- array(0L, dim(x))

  ## The workhorse; look in all four directions and add cells to the
  ## queue that are not blank and which have not been looked at yet.
  ## Mark any queued cell as looked at at each direction to avoid
  ## adding it four times and blowing the queue out.
  check <- function(q) {
    f <- function(d) {
      tmp <- q + rep(d, each=nrow(q))
      ret <- tmp[x[tmp] & valid[tmp], , drop=FALSE]
      valid[tmp] <<- FALSE
      ret
    }
    ret <- lapply(list(c(-1L, 0L), c(1L, 0L), c(0L, -1L), c(0L, 1L)), f)
    do.call("rbind", ret, quote=TRUE)
  }

  j <- 1L
  while (any(x)) {
    q <- which(x, TRUE, FALSE)[1L, , drop=FALSE]
    while (nrow(q) > 0L) {
      grp[q] <- j
      q <- check(q)
    }
    ## Square off the region that we found; assumes that nobody is
    ## embedding tables within holes in tables.  I think this is
    ## generally the correct thing based on the sheets that I have
    ## seen, but *someone* will be doing this, so perhaps changing
    ## this behaviour should be an option.  This will be necessary to
    ## rescue cells that are isolated in a table by being surrounded
    ## by blank cells.  An alternative approach would be not do this
    ## here, but do it at the end.  In that case we'd have to decide
    ## how to merge groups that overlap after squaring and applying
    ## additional heuristics to decide if they look like the same
    ## table.

    ## TODO: it's possible that the squaring here could allow for
    ## additional cells that need to be added to this group -- that's
    ## not checked at this point.  It's a nasty corner case but one
    ## that will turn up at some point.  We'd need to detect which
    ## cells are added at this point, re-queue them and see if they
    ## pick anything else up.  Most of the time they will not.
    r <- apply(which(grp == j, TRUE, FALSE), 2L, range)
    ir <- r[1L, 1L]:r[2L, 1L]
    ic <- r[1L, 2L]:r[2L, 2L]
    grp[ir, ic] <- j
    x[ir, ic] <- valid[ir, ic] <- FALSE
    j <- j + 1L
  }

  ## Drop the padding from 'grp'
  grp <- grp[seq_len(nr) + 1L, seq_len(nc) + 1L]
  ## writeLines(apply(grp, 1, paste, collapse=""))
  f <- function(idx) {
    r <- apply(which(grp == idx, TRUE, FALSE), 2, range)
    cellranger::cell_limits(r[1L, ], r[2L, ])
  }
  limits <- lapply(seq_len(j - 1L), f)

  switch(as,
         groups=grp,
         limits=limits,
         both=list(groups=grp, limits=limits))
}

##' @export
##' @rdname classify_sheet
split_sheet <- function(sheet) {
  if (!inherits(sheet, "worksheet")) {
    stop("sheet must be a 'worksheet' object")
  }
  limits <- classify_sheet(sheet, "limits")
  lapply(limits, worksheet_view, sheet=sheet)
}
