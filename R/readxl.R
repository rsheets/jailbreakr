##' Read an Excel spreadsheet the same way as readxl, but slower.
##' Assumes a well behaved table of data.
##' @title Read an Excel spreadsheet like readxl
##' @param path Path to the xlsx file
##' @param sheet Sheet name or an integer
##' @param col_names TRUE (the default) indicating we should use the
##'   first row as column names, FALSE, indicating we should generate
##'   names (X1, X2, ..., Xn) or a character vector of names to apply.
##' @param col_types Either NULL (the default) indicating we should
##'   guess the column types or a vector of column types (values must
##'   be "blank", "numeric", "date" or "text").
##' @param na Values indicating missing values (if different from
##'   blank).  Not yet used.
##' @param skip Number of rows to skip.
##' @export
jailbreak_readxl <- function(path, sheet=1L, col_names=TRUE,
                             col_types=NULL, na="", skip=0) {
  if (!identical(na, "")) {
    .NotYetUsed("na") # TODO
  }
  dat <- jailbreak_read(path, sheet)
  nc <- dat$dim[[2]]

  if (is.logical(col_names)) {
    if (col_names) {
      col_names <- rep(NA_character_, nc)
      i <- dat$lookup[skip + 1L, ]
      keep <- !is.na(i)
      col_names[keep] <- unlist(dat$cells$value[i[keep]])
      skip <- skip + 1L
    } else {
      col_names <- NULL
      keep <- rep(FALSE, nc)
    }
  } else if (is.character(col_names)) {
    if (length(col_names) != nc) {
      stop(sprintf("Expected %d col_names", nc))
    }
  } else {
    stop("`col_names` must be a logical or character vector")
  }

  lookup <- (if (skip > 0L) dat$lookup[-seq_len(skip), , drop=FALSE]
             else dat$lookup)

  if (!is.null(col_types)) {
    stop("Not yet handled")
    ## col_types: Either ‘NULL’ to guess from the spreadsheet or a
    ##           character vector containing "blank", "numeric",
    ##           "date" or "text".
  }

  ## Valid types: blank > date > numeric > text
  ## TODO: I can't do date yet.
  cell_types <- c(blank=0, date=1, number=2, text=3)
  type <- array(cell_types[["blank"]], dim(lookup))
  i <- na.omit(c(lookup))
  j <- !is.na(c(lookup))
  type[j][dat$cells$is_number[i]] <- cell_types[["number"]]
  type[j][dat$cells$is_text[i]] <- cell_types[["text"]]
  type <- apply(type, 2, max)
  type <- names(cell_types)[match(type, cell_types)]

  ret <- data.frame(array(NA, dim(lookup)))

  ## TODO: What becomes of boolean data from excel?
  tr <- list(number=as.numeric,
             date=function(...) stop("not implemented"),
             text=as.character,
             blank=as.numeric)
  for (i in seq_along(type)) {
    t <- type[[i]]
    j <- lookup[, i]
    k <- !is.na(j)
    ret[k, i] <- tr[[t]](dat$cells$value[j[k]])
  }

  keep <- keep | type != "blank"
  ret <- ret[keep]

  if (is.null(col_names)) {
    col_names <- sprintf("X%d", seq_len(ncol(ret)))
  } else {
    col_names <- col_names[keep]
  }
  names(ret) <- col_names

  class(ret) <- c("tbl_df", "tbl", "data.frame")
  ret
}
