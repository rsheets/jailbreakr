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
    .NotYetUsed("na") # TODO -- passed in to the type inference stuff
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

  ## Valid types: blank > bool > date > numeric > text
  ## NOTE: bool is going to map to map to number here to match the
  ## behaviour of readxl; I'd rather map it to logical, but I'm sure
  ## Hadley had a very good reason for not doing this, which I'd
  ## rather not find out the hard way.  At the same time I saw a
  ## comment in the source with a TODO on it --
  ##   src/XlsxCell.h: XlsxCell::type()
  ## TODO: columns that mix bool and number should throw an error at
  ## the moment because those types can't be harmonised, really.
  cell_types <- c(blank=0, bool=1, date=2, number=3, text=4)
  type <- array(unname(cell_types[dat$cells$type[c(lookup)]]), dim(lookup))
  type <- apply(rbind(cell_types[["blank"]], type), 2, max, na.rm=TRUE)
  type <- names(cell_types)[match(type, cell_types)]

  ret <- data.frame(array(NA, dim(lookup)))

  tr <- list(number=as.numeric,
             bool=as.numeric,
             date=unlist_times,
             text=as.character,
             blank=as.numeric)
  for (i in seq_along(type)) {
    t <- type[[i]]
    j <- lookup[, i]
    k <- !is.na(j)
    if (t == "date") {
      ## FFS. http://i.giphy.com/arz9UYo8bCo4E.gif
      ret[[i]] <- as.POSIXct(ret[[i]])
    }
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

## The R time objects really want me poke my eyes out.  Perhaps there
## is a better way of doing this?  Who knows?
unlist_times <- function(x) {
  if (length(x) == 0L) {
    structure(numeric(0), class=c("POSIXct", "POSIXt"), tzone="UTC")
  } else {
    tmp <- unlist(x)
    attributes(tmp) <- attributes(x[[1L]])
    tmp
  }
}
