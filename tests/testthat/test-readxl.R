## Test reading against readxl
context("readxl")

## TODO: I don't see the inline string version opening correctly in
## numbers so this might be a little beyond our needs to open here.
##
## TODO: I don't do date loading yet so am skipping the types sheet.
## Also handling of the boolean columns is incorrect, heading to
## character rather than logical.
test_that("agree with readxl", {
  files <- dir(file.path(get_readxl(), "tests/testthat"),
               "\\.xlsx$", full.names=TRUE)
  for (f in files) {
    if (grepl("^(inlineStr|types)", basename(f))) {
      next
    }
    message(f)
    cmp <- readxl::read_excel(f)
    dat <- jailbreak_readxl(f)
    expect_equal(dat, cmp)
  }
})
