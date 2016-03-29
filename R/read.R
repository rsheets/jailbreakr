##' This function does not get the data into a usable form but at
##' least loads it up into R so we can poke about with it.  Dates are
##' not yet handled because they're complicated.  The resulting loaded
##' data can distinguish between formulae and data, numbers and text.
##' Merged cells are detected.  A summary of the data is printed if
##' you print the resulting object.
##'
##' @title Read an xlsx file that probably contains nontabular data
##'
##' @param path Path to the xlsx file to load.  xls files are not supported.
##'
##' @param sheet Sheet number (not name at this point).  Googlesheets
##'   exported sheets are likely not to do the right thing.
##'
##' @return An \code{xlsx} object that can be printed.  Future methods
##'   might do something sensible.  The structure is subject to
##'   complete change so is not documented here.
##' @export
jailbreak_read <- function(path, sheet=1L) {
  ## TODO: This suffers from the same bug as readxl
  ##   htps://github.com/hadley/readxl/issues/104
  ## though it should be quite easy to fix in plain R; just need to
  ## resolve the relationship in _rels/workbook.xml.rels file if it
  ## exists (which my toy sheet does not have).
  ##
  ## NOTE: Some docs here:
  ##   https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.aspx
  ## though getting the actual spec would be nicer I suspect.
  if (!file.exists(path)) {
    stop(sprintf("%s does not exist", path))
  }
  xml <- xlsx_read_sheet(path, sheet)
  ns <- xml2::xml_ns(xml)
  strings <- xlsx_read_shared_strings(path)
  style <- xlsx_read_style(path)

  ## According to the spec mergeCells contains only mergeCell
  ## elements, and they contain only a "ref" attribute.  Once I track
  ## down the full schema (MS's website is a mess here) we can add
  ## correct references for this assertion.
  merged <- xml2::xml_text(
    xml2::xml_find_all(xml, "./d1:mergeCells/d1:mergeCell/@ref", ns))
  merged <- lapply(merged, cellranger::as.cell_limits)

  cells <- xlsx_parse_cells(xml, ns)
  cells_pos <- A1_to_matrix(cells$ref)

  ## I want to delete all merged cells from the cells list; forget
  ## about them as they inherit things from the anchor cell.
  if (length(merged) > 0L) {
    merged_pos <- lapply(merged, loc_merge, TRUE)
    merged_drop <- do.call("rbind", merged_pos)
    i <- match_cells(merged_drop, cells_pos)
    i <- -i[!is.na(i)]
    for (j in seq_along(cells)) {
      cells[[j]] <- cells[[j]][i]
    }
    cells_pos <- cells_pos[i, , drop=FALSE]
    tmp <- rbind(cells_pos, t(vapply(merged, function(el) el$lr, integer(2))))
    dim <- apply(tmp, 2, max)
  } else {
    dim <- apply(cells_pos, 2, max)
  }

  ## Now, build a look up table for all the cells.
  ## Lookup for "true" cells.
  lookup <- array(NA_integer_, dim)
  lookup[cells_pos] <- seq_len(nrow(cells_pos))

  ## A second table with merged cells, distinguished by being
  ## negative.  abs(lookup2) will give the correct value within the
  ## cells structure.
  if (length(merged) > 0L) {
    lookup2 <- lookup
    i <- match_cells(t(vapply(merged, function(x) x$ul, integer(2))), cells_pos)
    lookup2[merged_drop] <- -rep(i, vapply(merged_pos, nrow, integer(1)))
  } else {
    lookup2 <- lookup
  }

  ## Dealing with dates is a huge clustercuss; see the files
  ## CellType.h & XlsxWorkBook.h in readxl/src for a considerable
  ## amount of faffing to get dates working.  However, for my
  ## immediate needs we can juust skip over it.

  ## TODO: Roll this back into the xfs parsing perhaps?
  if ("formatCode" %in% names(style$num_formats)) {
    custom_date <- style$num_formats$numFmtId[
      grepl("[dmyhs]", style$num_formats$formatCode)]
  } else {
    custom_date <- integer()
  }
  ## Might roll this back into the style?
  is_date_time <- is_date_time(style$cell_xfs$numFmtId, custom_date)

  ## See readxl/src/XlsxCell.h: XlsxCell::type()

  ## String substitutions:
  i <- which(cells$type == "s")
  cells$value[i] <- strings[as.integer(unlist(cells$value[i])) + 1L]

  ## TODO: There are some blanks in here I need to get; formulae that
  ## yield zerolength strings, text cells that have no length.
  cells$is_formula <- lengths(cells$formula) > 0L
  cells$is_value <- lengths(cells$value) > 0L& !cells$is_formula

  type <- character(length(cells$value))
  type[!is.na(cells$type) & cells$type == "b"] <- "bool"
  type[!is.na(cells$type) & cells$type == "s" | cells$type == "str"] <- "text"
  i <- is.na(cells$type) | cells$type == "n"
  j <- is_date_time[cells$style[i] + 1L]
  type[i] <- ifelse(!is.na(j) & j, "date", "number")
  type[lengths(cells$value) == 0L] <- "blank"

  cells$type <- type
  cells$is_blank <- type == "blank"
  cells$is_bool <- type == "bool"
  cells$is_number <- cells$is_bool | type == "number"
  cells$is_date <- type == "date"
  cells$is_text <- type == "text"

  i <- cells$is_number
  cells$value[i] <- as.numeric(unlist(cells$value[i]))

  i <- cells$is_date
  cells$value[i] <-
    as.list(as.POSIXct(as.numeric(unlist(cells$value[i])) * 86400,
                       "UTC", xlsx_date_offset(path)))

  ret <- list(dim=dim, pos=cells_pos, cells=cells,
              merged=merged, style=style,
              lookup=lookup, lookup2=lookup2)
  class(ret) <- "xlsx"
  ret
}

##' @export
print.xlsx <- function(x, ...) {
  ## First, let's give an overview?
  dim <- x$dim
  cat(sprintf("<xlsx data: %d x %d>\n", dim[[1]], dim[[2]]))

  ## Helper for the merged cells.
  print_merge <- function(el) {
    anc <- "\U2693"
    left <- "\U2190"
    up <- "\U2191"
    ul <- "\U2196"

    d <- dim(el)
    anchor <- el$ul
    loc <- loc_merge(el)
    if (d[[1]] == 1L) {
      str <- rep(left, d[[2L]])
    } else if (d[[2L]] == 1L) {
      str <- rep(up, d[[1L]])
    } else {
      str <- matrix(ul, d[[1]], d[[2]])
      str[1L, ] <- left
      str[, 1L] <- up
    }
    str[[1L]] <- anc
    list(loc=loc, str=str)
  }

  m <- matrix(NA, dim[[1]], dim[[2]])
  for (el in x$merged) {
    tmp <- print_merge(el)
    m[tmp$loc] <- tmp$str
  }

  pos <- x$pos
  m[pos[x$cells$is_formula & x$cells$is_number, , drop=FALSE]] <- "="
  m[pos[x$cells$is_formula & x$cells$is_text,   , drop=FALSE]] <- "$"
  m[pos[x$cells$is_formula & x$cells$is_bool,   , drop=FALSE]] <- "!"
  m[pos[x$cells$is_formula & x$cells$is_date,   , drop=FALSE]] <- "#"
  m[pos[x$cells$is_value   & x$cells$is_number, , drop=FALSE]] <- "0"
  m[pos[x$cells$is_value   & x$cells$is_text,   , drop=FALSE]] <- "a"
  m[pos[x$cells$is_value   & x$cells$is_bool,   , drop=FALSE]] <- "b"
  m[pos[x$cells$is_formula & x$cells$is_date,   , drop=FALSE]] <- "d"
  m[is.na(m)] <- " "

  mm <- rbind(rep(LETTERS, length.out=dim[[2]]), m)
  cat(paste(sprintf("%s: %s\n",
                    format(c("", seq_len(dim[[1]]))),
                    apply(mm, 1, paste, collapse="")), collapse=""))
  invisible(x)
}

##' Subset a jailbreakr sheet based on a rectangular selection based
##' on a cellranger \code{cell_limits} object.
##'
##' @title Subset a jailbreakr sheet
##' @param dat A jailbreakr sheet
##' @param xr A \code{cell_limits} object indicating the extent to
##'   extract.
##' @export
jailbreak_subset <- function(dat, xr) {
  if (!inherits(xr, "cell_limits")) {
    stop("xr must be a cell_limits object")
  }

  x <- dat
  x$dim <- dim(xr)
  i <- (dat$pos[, 1L] >= xr$ul[[1L]] &
        dat$pos[, 1L] <= xr$lr[[1L]] &
        dat$pos[, 2L] >= xr$ul[[2L]] &
        dat$pos[, 2L] <= xr$lr[[2L]])

  x$cells <- lapply(dat$cells, function(x) x[i])

  ## Need to recompute the value here; rows and columns have both
  ## possibly shifted.
  x$pos <- dat$pos[i, , drop=FALSE] - rep(xr$ul - 1L, each=sum(i))

  ## The lookups need recomputing.  This is potentially quite nasty to
  ## do.
  j <- which(i)
  ir <- xr$ul[[1L]] : xr$lr[[1L]]
  ic <- xr$ul[[2L]] : xr$lr[[2L]]
  x$lookup <- dat$lookup[ir, ic]
  x$lookup2 <- dat$lookup2[ir, ic]
  x$lookup[] <- match(x$lookup, j)
  x$lookup2[] <- match(abs(x$lookup2), j) * sign(x$lookup2)

  ## Merged comes next.
  f <- function(el) {
    ok <- all(el$ul >= xr$ul & el$lr <= xr$lr)
    if (ok) {
      el$ul <- el$ul - xr$ul + c(1L, 1L)
      el$lr <- el$lr - xr$ul + c(1L, 1L)
      el
    } else {
      NULL
    }
  }
  merged <- lapply(x$merged, f)
  x$merged <- merged[!vlapply(merged, is.null)]

  x
}

## Non api functions:
## xlsx_read_*: reads something from a file
## xlsx_parse_*: turns xml into somethig usable

xlsx_read_sheet <- function(path, sheet) {
  xml <- xlsx_read_file(path, internal_sheet_name(path, sheet))
  stopifnot(xml2::xml_name(xml) == "worksheet")
  xml
}

xlsx_read_file <- function(path, file) {
  tmp <- tempfile()
  dir.create(tmp)
  ## Oh boy more terrible default behaviour.
  filename <- tryCatch(utils::unzip(path, file, exdir=tmp),
                       warning=function(e) stop(e))
  on.exit(unlink(tmp, recursive=TRUE))
  xml2::read_xml(filename)
}

xlsx_read_file_if_exists <- function(path, file, missing=NULL) {
  ## TODO: Appropriate error handling here is difficult; we should
  ## check that `path` exists, but by the time that this is called we
  ## know that already.
  tmp <- tempfile()
  dir.create(tmp)
  filename <- tryCatch(utils::unzip(path, file, exdir=tmp),
                       warning=function(e) NULL,
                       error=function(e) NULL)
  if (is.null(filename)) {
    missing
  } else {
    on.exit(unlink(tmp, recursive=TRUE))
    xml2::read_xml(filename)
  }
}

## If the format is <si>/<t> then we can just take the text values.
## Otherwise we'll have to parse out the RTF strings separately.
xlsx_read_shared_strings <- function(path) {
  xml <- xlsx_read_file_if_exists(path, "xl/sharedStrings.xml")
  if (is.null(xml)) {
    return(character(0))
  }
  ns <- xml2::xml_ns(xml)
  si <- xml2::xml_find_all(xml, "d1:si", ns)
  if (length(si) == 0L) {
    return(character(0))
  }
  ## TODO: This is a bug in xls
  is_rich <- xml2::xml_find_lgl(si, "boolean(d1:r)", ns)
  ret <- character(length(si))
  ret[!is_rich] <-
    xml2::xml_text(xml2::xml_find_one(si[!is_rich], "d1:t", ns))
  ret[is_rich] <- vcapply(si[is_rich], xlsx_parse_string, ns)
  ret
}

## sheetData: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx
##
##   Nothing looks interesting in sheetData, and all elements must be
##   'row'.
##
## row: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.row.aspx
##   The only interesting attribute here is "hidden", but that
##   includes being collapsed by outline, so probably not that
##   interesting.
##
## cell: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.cell.aspx
##
##   Might contain
##     <f>: formula
##     <is> rich test inline
##     <v> value
##   Interesting attributes:
##     r: an A1 style reference to the locatiopn of this cell
##     s: the index of this cell's style (if colours are used as a guide)
##     t: type "an enumeration representing the cell's data type", the
##       only reference to which I can find is
##       http://mailman.vse.cz/pipermail/sc34wg4/attachments/20100428/3fc0a446/attachment-0001.pdf
##       - b: boolean
##       - d: date (ISO 8601)
##       -  e: error
##       - inlineStr: inline string in rich text format, with
##           contents in the 'is' element, not the 'v' element.
##       - n: number
##       - s: shared string
##       - str: formula string
##
## However, many numbers seem not to have a "t" attribute which is
## charming.
##
## NOTE: handling of formulae is potentially tricky as they can have an attribute "shared" which
##
## Blank cells have no children at all.
xlsx_parse_cells <- function(xml, ns) {
  sheet_data <- xml2::xml_find_one(xml, "d1:sheetData", ns)
  cells <- xml2::xml_find_all(sheet_data, "./d1:row/d1:c", ns)

  xml_find_if_exists <- function(x, xpath, ns) {
    i <- xml2::xml_find_lgl(x, sprintf("boolean(%s)", xpath), ns)
    ret <- vector("list", length(i))
    ret[i] <- xml2::xml_text(xml2::xml_find_one(x[i], xpath, ns))
    ret
  }

  ref <- xml2::xml_attr(cells, "r")
  style <- as.integer(xml2::xml_attr(cells, "s"))
  type <- xml2::xml_attr(cells, "t")

  formula <- xml_find_if_exists(cells, "./d1:f", ns)
  value <- xml_find_if_exists(cells, "./d1:v", ns)

  ## Quick check to make sure we didn't miss anything (I think it's
  ## only is values)
  inline_string <- xml2::xml_find_lgl(cells, "boolean(./d1:is)", ns)
  if (any(inline_string)) {
    ## These would get fired through the string parsing I think.
    stop("Inline string value not yet handled")
  }

  list(ref=ref, style=style, type=type,
       formula=formula, value=value)
}

xlsx_parse_string <- function(x, ns) {
  t <- xml2::xml_find_one(x, "d1:t", ns)
  if (inherits(t, "xml_missing")) {
    ## NOTE: it *looks* like most of the time we can do xml_text(x)
    ## here and get about the right answer.
    str <- character()
  } else {
    str <- xml2::xml_text(t)
  }
  r <- xml2::xml_find_all(x, "d1:r", ns)
  if (length(r) > 0L) {
    str <- paste(c(str, xml2::xml_text(r)), collapse="")
  }
  ## TODO: still need to "unescape" these.
  str
}

loc_merge <- function(el, drop_anchor=FALSE) {
  d <- dim(el)
  anchor <- el$ul
  if (d[[1]] == 1L) {
    rows <- anchor[[1]]
    cols <- seq.int(anchor[[2]], by=1L, length.out=d[[2L]])
  } else if (d[[2L]] == 1L) {
    rows <- seq.int(anchor[[1]], by=1L, length.out=d[[1L]])
    cols <- anchor[[2]]
  } else {
    cols <- seq.int(anchor[[2]], by=1L, length.out=d[[2L]])
    rows <- seq.int(anchor[[1]], by=1L, length.out=d[[1L]])
  }
  ret <- cbind(row=rows, col=cols)
  if (drop_anchor) {
    ret[-1, , drop=FALSE]
  } else {
    ret
  }
}

match_cells <- function(x, table, ...) {
  ## assumes 2-column integer matrix
  x <- paste(x[, 1L], x[, 2L], sep="\r")
  table <- paste(table[, 1L], table[, 2L], sep="\r")
  match(x, table, ...)
}

sheet_names <- function(filename) {
  xml <- xlsx_read_file(filename, "xl/workbook.xml")
  ns <- xml2::xml_ns(xml)
  xml2::xml_text(xml2::xml_find_all(xml, "d1:sheets/d1:sheet/@name", ns))
}

## Return the filename within the bundle
internal_sheet_name <- function(filename, sheet) {
  if (length(sheet) != 1L) {
    stop("'sheet' must be a scalar")
  }
  if (is.na(sheet)) {
    stop("'sheet' must be non-missing")
  }
  if (is.character(sheet)) {
    sheet <- match(sheet, sheet_names(filename))
  } else if (!(is.integer(sheet) || is.numeric(sheet))) {
    stop("'sheet' must be an integer or a string")
  }

  ## TODO: Looks like this does always exist.
  rels <- xlsx_read_file_if_exists(filename, "xl/_rels/workbook.xml.rels")
  if (is.null(rels)) {
    target <- sprintf("xl/worksheets/sheet%d.xml", sheet)
  } else {
    ## This might fail with a cryptic error if my assumptions are
    ## incorrect.
    xml <- xlsx_read_file(filename, "xl/workbook.xml")
    xpath <- sprintf("d1:sheets/d1:sheet[%d]", sheet)
    node <- xml2::xml_find_one(xml, xpath, xml2::xml_ns(xml))
    id <- xml2::xml_attr(node, "id")
    ## This _should_ work but I don't see it:
    ##   xpath <- sprintf("string(d1:sheets/d1:sheet[%d]/@id)", sheet)
    ##   xml2::xml_find_chr(xml, xpath, ns) # --> ""
    xpath <- sprintf('/d1:Relationships/d1:Relationship[@Id = "%s"]/@Target',
                     id)
    target <- xml2::xml_text(xml2::xml_find_one(rels, xpath,
                                                xml2::xml_ns(rels)))
    ## NOTE: these are _relative_ paths so must be qualified here:
    target <- file.path("xl", target)
  }
  target
}

xlsx_date_offset <- function(path) {
  ## See readxl/src/utils.h: dateOffset
  ## See readxl/src/XlsxWorkbook.h: is1904
  xml <- xlsx_read_file(path, "xl/workbook.xml")
  date1904 <- xml2::xml_find_one(xml, "/d1:workbook/d1:workbookPr/@date1904",
                                 xml2::xml_ns(xml))
  if (inherits(date1904, "xml_missing")) {
    date_is_1904 <- FALSE
  } else {
    ## TODO: in theory we should do whatever atoi would allow here
    ## (that's what Hadley uses in the C++) but I have a sheet that
    ## contains this as "false".  So I'm trying this way for now.
    value <- xml2::xml_text(date1904)
    date_is_1904 <- value == "1" || value == "true"
  }
  if (date_is_1904) "1904-01-01" else "1899-12-30"
}

is_date_time <- function(id, custom) {
  ## See readxl's src/CellType.h: isDateTime()
  id %in% c(c(14:22, 27:36, 45:47, 50:58, 71:81), custom)
}
