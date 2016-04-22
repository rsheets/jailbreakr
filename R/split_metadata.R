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
##' @param sheet A worksheet or a worksheet view
##' @return A worksheet view
##' @author Rich FitzJohn
split_metadata <- function(x, include_merged=TRUE, min_jump=2,
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

  ans <- detect_metadata(!is.na(lookup), min_jump, min_data_block)

  if (inherits(x, "worksheet_view")) {
    xr <- x$xr
    data <- xr$data
  } else {
    xr <- cellranger::cell_limits(c(1, 1), sheet$dim)
    data <- list()
  }
  xr_meta <- xr
  xr$ul[[1]] <- xr$ul[[1]] + ans

  xr_meta$lr[[1]] <- xr_meta$ul[[1]] + ans - 1L
  data$metadata <- linen::worksheet_view(sheet, xr_meta)

  linen::worksheet_view(sheet, xr, data)
}

detect_metadata <- function(m, min_jump=2, min_data_block=5) {
  i <- apply(m, 1, function(x) max(which(x)))

  ## Now, we look for the first increase of size min_jump where there
  ## is not a further jump (of any size) for min_data_block.  This is
  ## not going to deal terribly well with missing data in the far
  ## right column which will cause the number to move in and out.

  ## These are all the increases; these are the only things to deal with?
  di <- diff(i)
  j <- which(di > 0L)
  for (k in j) {
    if (di[[k]] >= min_jump) {
      if (max(i[(k + 1L):min(k + 1L + min_data_block, length(i))]) -
          i[[k + 1L]] == 0) {
        return(k)
      }
    }
  }
  0L
}
