vlapply <- function(X, FUN, ...) {
  vapply(X, FUN, logical(1), ...)
}
viapply <- function(X, FUN, ...) {
  vapply(X, FUN, integer(1), ...)
}
vnapply <- function(X, FUN, ...) {
  vapply(X, FUN, numeric(1), ...)
}
vcapply <- function(X, FUN, ...) {
  vapply(X, FUN, character(1), ...)
}

## Could contribute to cellranger?
A1_to_matrix <- function(x) {
  char0_to_NA <- cellranger:::char0_to_NA
  rm_dollar_signs <- cellranger:::rm_dollar_signs

  stopifnot(is.character(x))

  x <- rm_dollar_signs(x)

  m <- regexec("[[:digit:]]*$", x)
  m <- regmatches(x, m)
  row_part <- as.integer(vapply(m, char0_to_NA, character(1)))
  row_part <- ifelse(row_part > 0, row_part, NA)

  m <- regexec("^[[:alpha:]]*", x)
  m <- regmatches(x, m)
  col_part <- cellranger::letter_to_num(vapply(m, char0_to_NA, character(1)))

  cbind(row=row_part, col=col_part)
}
