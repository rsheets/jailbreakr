context("googlesheets")

test_that("can read google sheets in right order", {
  skip_if_not_installed("googlesheets")
  filename <- system.file("mini-gap.xlsx", package="googlesheets")

  expect_equal(internal_sheet_name(filename, 1L),
               "xl/worksheets/sheet4.xml")
})
