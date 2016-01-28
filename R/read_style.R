xlsx_read_style <- function(path) {
  xml <- xlsx_read_file(path, "xl/styles.xml")
  ns <- xml2::xml_ns(xml)

  fonts <- xlsx_read_style_fonts(xml, ns)
  fills <- xlsx_read_style_fills(xml, ns)
  borders <- xlsx_read_style_borders(xml, ns)

  list(fonts=fonts, fills=fills, borders=borders)
}
xlsx_read_style_fonts <- function(xml, ns) {
  fonts <- xml2::xml_find_one(xml, "d1:fonts", ns)
  fonts_f <- xml2::xml_children(fonts)

  pos <- sort(unique(unlist(lapply(fonts_f, function(el)
    xml2::xml_name(kids <- xml2::xml_children(el))))))
  known <- c("sz", "color", "scheme", "family", "b", "i", "u", "name")
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
    ret <- xml2::xml_attrs(x)
    nms <- sort(unique(unlist(lapply(ret, names))))
    ret <- t(vapply(ret, function(x) x[nms], character(length(nms))))
    if (length(nms) == 1L) {
      ret <- t(ret)
    }
    if (length(nms) > 0L) {
      colnames(ret) <- paste0(prefix, nms)
    }
    data.frame(ret, stringsAsFactors=FALSE)
  }

  cbind(type=type, f(fg, "fg_"), f(bg, "bg_"), stringsAsFactors=FALSE)
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
