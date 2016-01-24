context("read")

test_that("simple", {
  path <- system.file("extdata/example.xlsx", package="jailbreakr")
  dat <- jailbreak_read(path)
  expect_is(dat, "xlsx")
  expect_equal_to_reference(capture.output(print(dat)),
                            "reference/example.xlsx_print.rds")
})
