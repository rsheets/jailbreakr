xlsx_read_style <- function(path) {
  xml <- xlsx_read_file(path, "xl/styles.xml")
  ns <- xml2::xml_ns(xml)

  fonts <- xlsx_read_style_fonts(xml, ns)
  fills <- xlsx_read_style_fills(xml, ns)
  borders <- xlsx_read_style_borders(xml, ns)

  ## XFS is "cell formatting".  The s="<int>" tag refers to an entry
  ## in cellXfs, so this is _probably_ the most useful.
  cell_style_xfs <- xlsx_read_cell_style_xfs(xml, ns)
  cell_xfs <- xlsx_read_cell_xfs(xml, ns)
  cell_styles <- xlsx_read_cell_styles(xml, ns)
  num_formats <- xlsx_read_num_formats(xml, ns)

  list(fonts=fonts, fills=fills, borders=borders,
       cell_style_xfs=cell_style_xfs,
       cell_xfs=cell_xfs,
       cell_styles=cell_styles,
       num_formats=num_formats)
}

xlsx_read_style_fonts <- function(xml, ns) {
  fonts <- xml2::xml_find_one(xml, "d1:fonts", ns)
  fonts_f <- xml2::xml_children(fonts)

  pos <- sort(unique(unlist(lapply(fonts_f, function(el)
    xml2::xml_name(kids <- xml2::xml_children(el))))))
  known <- c("sz", "color", "scheme", "family", "b", "i", "u", "name",
             ## These we ignore:
             "charset")
  unk <- setdiff(pos, known)
  if (length(unk) > 0L) {
    message("Skipping unhandled font tags: ", paste(unk, collapse=", "))
  }

  sz <- xml2::xml_text(xml2::xml_find_one(fonts_f, "d1:sz/@val", ns))
  col <- xml2::xml_text(xml2::xml_find_one(fonts_f, "d1:color/@theme", ns))
  scheme <- xml2::xml_text(xml2::xml_find_one(fonts_f, "d1:scheme/@val", ns))
  family <- xml2::xml_text(xml2::xml_find_one(fonts_f, "d1:family/@val", ns))
  name <- xml2::xml_text(xml2::xml_find_one(fonts_f, "d1:name/@val", ns))
  b <- xml2::xml_find_lgl(fonts_f, "boolean(d1:b)", ns)
  i <- xml2::xml_find_lgl(fonts_f, "boolean(d1:i)", ns)
  u <- xml2::xml_find_lgl(fonts_f, "boolean(d1:u)", ns)

  data.frame(name=name, family=family, scheme=scheme, col=col,
             size=as.integer(sz), bold=b, italic=i, underline=u,
             stringsAsFactors=FALSE)
}

xlsx_read_style_fills <- function(xml, ns) {
  fills <- xml2::xml_find_one(xml, "d1:fills", ns)
  fills_f <- xml2::xml_find_one(xml2::xml_children(fills), "d1:patternFill", ns)
  missing <- vlapply(fills_f, inherits, "xml_missing")
  if (any(missing)) {
    stop("I can only handle patternFill at the moment")
  }

  type <- xml2::xml_attr(fills_f, "patternType")
  fg <- xml2::xml_find_one(fills_f, "d1:fgColor", ns)
  bg <- xml2::xml_find_one(fills_f, "d1:bgColor", ns)

  ## There's some crap to deal here with themes and tints.  It's not
  ## totally clear to me which things are valid for what because I see
  ## only:
  ##   fg: theme, tiny, rgb
  ##   bg: indexed
  f <- function(x, prefix) {
    ret <- data.frame(attrs_to_matrix(x), stringsAsFactors=FALSE)
    names(ret) <- paste0(prefix, names(ret))
    ret
  }

  ret <- data.frame(type=type, stringsAsFactors=FALSE)
  if (!all(vlapply(fg, inherits, "xml_missing"))) {
    ret <- cbind(ret, f(fg, "fg_"))
  }
  if (!all(vlapply(bg, inherits, "xml_missing"))) {
    ret <- cbind(ret, f(bg, "bg_"))
  }
  ret
}

xlsx_read_style_borders <- function(xml, ns) {
  borders <- xml2::xml_find_one(xml, "d1:borders", ns)
  borders_f <- xml2::xml_children(borders)

  pos <- sort(unique(unlist(lapply(borders_f, function(el)
    xml2::xml_name(kids <- xml2::xml_children(el))))))
  known <- c("bottom", "left", "top", "right", "diagonal")
  unk <- setdiff(pos, known)
  if (length(unk) > 0L) {
    message("Skipping unhandled border tags: ", paste(unk, collapse=", "))
  }

  ## NOTE: Not taking any actual style information here (e.g., weight,
  ## colour), aside from the presence of the border.  It's possible
  ## that this does not get the correct style in all cases though --
  ## if style == "none" perhaps this is wrong?  Is that a valid style?
  data.frame(
    bottom=xml2::xml_find_lgl(borders_f, "boolean(d1:bottom/@style)", ns),
    left=xml2::xml_find_lgl(borders_f, "boolean(d1:left/@style)", ns),
    top=xml2::xml_find_lgl(borders_f, "boolean(d1:top/@style)", ns),
    right=xml2::xml_find_lgl(borders_f, "boolean(d1:right/@style)", ns),
    diagonal=xml2::xml_find_lgl(borders_f, "boolean(d1:diagonal/@style)", ns),
    stringsAsFactors=FALSE)
}

xlsx_read_cell_style_xfs <- function(xml, ns) {
  csx <- xml2::xml_find_one(xml, "d1:cellStyleXfs", ns)
  as.data.frame(attrs_to_matrix(xml2::xml_children(csx), "integer"))
}

xlsx_read_cell_xfs <- function(xml, ns) {
  cx <- xml2::xml_find_one(xml, "d1:cellXfs", ns)
  cx_kids <- xml2::xml_children(cx)
  ret <- as.data.frame(attrs_to_matrix(cx_kids, "integer"))
  ret_align <- attrs_to_matrix(xml2::xml_find_one(cx_kids, "d1:alignment", ns))
  ret_align <- data.frame(ret_align, stringsAsFactors=FALSE)
  if ("wrapText" %in% names(ret_align)) {
    ret_align$wrapText <- as.logical(ret_align$wrapText)
  }
  cbind(ret, ret_align)
}

xlsx_read_cell_styles <- function(xml, ns) {
  cs <- xml2::xml_find_one(xml, "d1:cellStyles", ns)
  ret <- data.frame(attrs_to_matrix(xml2::xml_children(cs)),
                    stringsAsFactors=FALSE)
  if ("xfId" %in% names(ret)) {
    ret$xfId <- as.integer(ret$xfId)
  }
  if ("builtinId" %in% names(ret)) {
    ret$builtinId <- as.integer(ret$builtinId)
  }
  if ("hidden" %in% names(ret)) {
    ret$hidden <- as.logical(ret$hidden)
  }
  ret
}

xlsx_read_num_formats <- function(xml, ns) {
  dat <- xml2::xml_find_one(xml, "d1:numFmts", ns)
  ret <- as.data.frame(attrs_to_matrix(xml2::xml_children(dat)))
  if ("numFmtId" %in% names(ret)) {
    ret$numFmtId <- as.integer(ret$numFmtId)
  }
  ret
}

attrs_to_matrix <- function(x, mode=NULL) {
  dat <- xml2::xml_attrs(x)
  nms <- unique(unlist(lapply(dat, names)))
  ret <- t(vapply(dat, function(x) x[nms], character(length(nms))))
  if (length(nms) == 1L) {
    ret <- t(ret)
  }
  if (length(nms) > 0L) {
    colnames(ret) <- nms
  }
  if (!is.null(mode)) {
    storage.mode(ret) <- mode
  }
  ret
}
