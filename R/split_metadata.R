##' Split metadata from a worksheet or region of a worksheet (a
##' worksheet view).
##'
##' There is a common pattern in spreadsheets where data is stored as:
##'
##' \preformatted{
##' mmm
##' mmm
##' HHHHHHHH
##' dddddddd
##' dddddddd
##' }
##'
##' where \code{m} is some metadata (perhaps indicating table name,
##' creator, dates, etc), \code{H} is the header and \code{d} is the
##' actual data.  This function will split the metadata (\code{m})
##' part off, leaving a table that is more suitable for further
##' processing.  In many ways this is like the \code{skip} argument to
##' \code{readxl::read_excel} and \code{read.csv}, but we will retain
##' the metadata (somewhere!).
##'
##' The idea here is that the metadata block starts when we get a
##' shift in the number of non-blank cells.  There needs to be some
##' heuristics here to help, and things will need to be tuneable: do
##' merged cells count as non-empty (\code{include_merged}; default is
##' to include them), how big a jump we look for (\code{min_jump};
##' default is 2 columns), how many rows of the same size do we look
##' for in the data block (\code{min_data_block}; default is 5 rows).
##'
##' Other things that might be useful, but which aren't supported yet,
##' include looking for different colours and fonts in the metadata
##' and the main block; when we switch from one to the other we're
##' likely to see things like a change here.
##'
##' @title Split metadata from a worksheet
##' @param x A worksheet or a worksheet view
##' @param include_merged Include merged cells when computing row
##'   width.  This is generally going to be what you want to do, but
##'   if you have a sheet with metadata rows that are entirely merged
##'   across, you probably will do better by turning this option off
##'   (in that case they'd be counted as a single column wide).
##' @param min_jump The minimum number of columns increase that we'll
##'   count as moving from the metadata block to the data block.  A
##'   single column (1) is probably going to be prone to false
##'   positives, and in sheets with this pattern the jump is often
##'   quite large.
##' @param min_data_block The minimum number of rows without
##'   decreasing in size before we can conclude that we're in the data
##'   block.  If we run off the end of the worksheet before reaching
##'   this number, we'll conclude no metadata was found.
##' @return For \code{split_metadata} and \code{split_metadata_apply},
##'   a worksheet view; in this view the \code{data$metadata} element
##'   will be a worksheet view of the metadata.  For
##'   \code{split_metadata_find}, a single integer representing the
##'   number of rows of metadata found (with zero indicating no
##'   metadata).
##' @export
split_metadata <- function(x, include_merged=TRUE, min_jump=2,
                           min_data_block=5) {
  n <- split_metadata_find(x, include_merged, min_jump, min_data_block)
  split_metadata_apply(x, n)
}

##' @export
##' @rdname split_metadata
split_metadata_find <- function(x, include_merged=TRUE, min_jump=2,
                                min_data_block=5) {
  ## This is general thing here, I think.
  if (inherits(x, "worksheet_view")) {
    lookup <- x$sheet$lookup2[x$idx$r, x$idx$c]
    sheet <- x$sheet
  } else {
    lookup <- x$lookup2
    sheet <- x
  }

  lookup[sheet$cells$type[abs(lookup)] == "blank"] <- NA
  if (!include_merged) {
    lookup[lookup < 0] <- NA
  }

  split_metadata_classify(!is.na(lookup), min_jump, min_data_block)
}

##' @export
##' @rdname split_metadata
##' @param n A scalar integer representing the number of rows to
##'   consider to be metadata (i.e., equivalent to the return value of
##'   \code{split_metadata_find}.
split_metadata_apply <- function(x, n) {
  if (inherits(x, "worksheet_view")) {
    xr <- x$xr
    data <- xr$data
  } else {
    xr <- cellranger::cell_limits(c(1, 1), sheet$dim)
    data <- list()
  }
  xr_meta <- xr
  xr$ul[[1]] <- xr$ul[[1]] + n

  xr_meta$lr[[1]] <- xr_meta$ul[[1]] + n - 1L
  data$metadata <- linen::worksheet_view(sheet, xr_meta)

  linen::worksheet_view(sheet, xr, data)
}

## Now, we look for the first increase of size min_jump where there is
## not a further jump (of any size) for min_data_block.  This is not
## going to deal terribly well with missing data in the far right
## column which will cause the number to move in and out.
split_metadata_classify <- function(m, min_jump, min_data_block) {
  i <- apply(m, 1, function(x) max(which(x)))
  di <- diff(i)
  j <- which(di > 0L)
  for (k in j) {
    if (di[[k]] >= min_jump) {
      n <- k + min_data_block
      if (n > length(i)) {
        break
      } else if (all(i[(k + 1L):n] <= i[[k + 1L]])) { # no growth
        return(k)
      }
    }
  }
  0L
}
