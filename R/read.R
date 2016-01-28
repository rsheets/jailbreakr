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
  strings <- xlsx_read_strings(path)
  ## NOTE: not currently used; still processing this.
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
  merged_pos <- lapply(merged, loc_merge, TRUE)
  merged_drop <- do.call("rbind", merged_pos)
  i <- match_cells(merged_drop, cells_pos)
  i <- -i[!is.na(i)]
  for (j in seq_along(cells)) {
    cells[[j]] <- cells[[j]][i]
  }
  pos <- cells_pos[i, , drop=FALSE]

  ## Now, build a look up table for all the cells.

  ## NOTE: pos is not actually enough because it won't take into
  ## account any merged cells if they flow out from the sides of the
  ## data, and some *will* do that, given enough sheets.
  tmp <- rbind(pos, t(vapply(merged, function(el) el$lr, integer(2))))
  dim <- apply(tmp, 2, max)

  ## Lookup for "true" cells.
  lookup <- array(NA_integer_, dim)
  lookup[pos] <- seq_len(nrow(pos))
  ## A second table with merged cells, distinguished by being
  ## negative.  abs(lookup2) will give the correct value within the
  ## cells structure.
  lookup2 <- lookup
  i <- match_cells(t(vapply(merged, function(x) x$ul, integer(2))), pos)
  lookup2[merged_drop] <- -rep(i, vapply(merged_pos, nrow, integer(1)))

  ## Dealing with dates is a huge clustercuss; see the files
  ## CellType.h & XlsxWorkBook.h in readxl/src for a considerable
  ## amount of faffing to get dates working.  However, for my
  ## immediate needs we can juust skip over it.
  ##
  ## It's likely that formatting will give good hints about where
  ## logical regions in spreadsheets are.  In that case, see here:
  ## http://blogs.msdn.com/b/brian_jones/archive/2007/05/29/simple-spreadsheetml-file-part-3-formatting.aspx
  ## coloured blocks (fill style) and gridlines are probably going to
  ## be the trick.  Gridlines are going to be nasty though because of
  ## corners and that spreadsheets, as they age, get corners all over
  ## the show where they should not be.  So another thing for later.

  ## String substitutions:
  i <- which(cells$type == "s")
  cells$value[i] <- strings[as.integer(unlist(cells$value[i])) + 1L]
  i <- which(lengths(cells$value) > 0L & is.na(cells$type))
  cells$value[i] <- as.numeric(unlist(cells$value[i]))

  cells$is_formula <- lengths(cells$formula) > 0L
  cells$is_blank <- lengths(cells$value) == 0L
  cells$is_value <- !cells$is_blank & !cells$is_formula
  cl <- vcapply(cells$value, class)
  cells$is_number <- cl == "numeric"
  cells$is_text <- cl == "character"

  ret <- list(dim=dim, pos=pos, cells=cells,
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
  m[pos[x$cells$is_value   & x$cells$is_number, , drop=FALSE]] <- "0"
  m[pos[x$cells$is_value   & x$cells$is_text,   , drop=FALSE]] <- "a"
  m[is.na(m)] <- " "

  mm <- rbind(rep(LETTERS, length.out=dim[[2]]), m)
  cat(paste(sprintf("%s: %s\n",
                    format(c("", seq_len(dim[[1]]))),
                    apply(mm, 1, paste, collapse="")), collapse=""))
  invisible(x)
}

## Non api functions:
## xlsx_read_*: reads something from a file
## xlsx_parse_*: turns xml into somethig usable

xlsx_read_sheet <- function(path, sheet) {
  xml <- xlsx_read_file(path, sprintf("xl/worksheets/sheet%d.xml", sheet))
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

xlsx_read_strings <- function(path) {
  xml <- xlsx_read_file(path, "xl/sharedStrings.xml")
  ns <- xml2::xml_ns(xml)
  xlsx_parse_strings(xml2::xml_find_all(xml, "d1:si", ns), ns)
}

xlsx_find_contents <- function(x, ..., text = TRUE) {
  value <- tryCatch(xml2::xml_find_one(x, ...), error = function(e) NULL)
  if (text && !is.null(value)) {
    value <- xml2::xml_text(value)
  }
  value
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
  style <- xml2::xml_attr(cells, "s")
  type <- xml2::xml_attr(cells, "t")

  formula <- xml_find_if_exists(cells, "./d1:f", ns)
  value <- xml_find_if_exists(cells, "./d1:v", ns)

  ## Quick check to make sure we didn't miss anything (I think it's
  ## only is values)
  if (any(xml2::xml_find_lgl(cells, "boolean(./d1:is)", ns))) {
    stop("Inline string value not yet handled")
  }

  list(ref=ref, style=style, type=type,
       formula=formula, value=value)
}

## If the format is <si>/<t> then we can just take the text values.
## Otherwise we'll have to parse out the RTF strings separately.
xlsx_parse_strings <- function(x, ns) {
  res <- xml2::xml_find_all(x, "./d1:si/d1:t", ns)
  if (length(res) == length(x)) {
    as.list(xml2::xml_text(res))
  } else {
    lapply(x, xlsx_parse_string, ns)
  }
}

## TODO: If this is the major timesink, then could go ahead and be one
## xpath by looking for si/t?
##
## This could be sped up by a factor of 10 if we take vectorised input.
xlsx_parse_string <- function(x, ns) {
  t <- tryCatch(xml2::xml_find_one(x, "d1:t", ns), error=function(e) NULL)
  if (is.null(t)) {
    str <- character()
  } else {
    str <- xml2::xml_text(x)
  }

  r <- xml2::xml_find_all(x, "d1:r", ns)
  nr <- length(r)
  if (nr > 0L) {
    str2 <- character(nr)
    found <- logical(nr)
    for (i in seq_len(nr)) {
      t <- tryCatch(xml2::xml_find_one(r[[i]], "d1:t", ns),
                    error=function(e) NULL)
      found[[i]] <- !is.null(t)
      if (found[[i]]) {
        str2[[i]] <- as.character(xml2::xml_contents(t))
      }
    }
    str <- c(str, str2[found])
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
