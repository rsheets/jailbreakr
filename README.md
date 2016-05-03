# jailbreakr

**Warning: This project is in the early scoping stages; do not use for anything other than amusement/frustration purposes**

Data Liberator.  To extract tabular data people put in nontabular structures in a program designed to hold tables.

![](http://i.giphy.com/SEp6Zq6ZkzUNW.gif)

## Installation

Requires the development version of xml2 (for `xml_find_lgl`)

```r
devtools::install_github("rsheets/jailbreakr")
```

## Goals

There are two large excel spreadsheet corpuses; it would be nice to use these to get a feel for what fraction of spreadsheets we can handle or the range of non-table-like data out there.

![the things people do to data](http://replygif.net/i/514.gif)

The first is the [EUSES corpus](http://openscience.us/repo/spreadsheet/euses.html) of 4,447 spreadsheets (16,853 worksheets).  This is all xls files (rather than xlsx) and therefore neaed either an [xls -> xlsx conversion](http://bit.ly/1P2rMGr) or support in jailbreakr for xls files.

The second, larger, one is the [Enron corpus](http://www.felienne.com/archives/3634) of 15,770 spreadsheets (79,983)

# Roadmap

* data structure package:
  - linen?  General representation of spreadsheet data, plus some limited low-level operations on that data
  - depends on cell ranger, tibble
  - constructor function
  - print methods
  - subsetting, range extraction etc.
  - plot method - for quickly getting a feel for structure, or a shiny app
  - summary: this has n sheets, no formulae, 3 plots, etc, things about the references between the sheets?
  - where it came from (excel, googlesheet, etc), with filenames, reference ids etc.
  - probably needs references to handle multiple sheets and formulae within them, definitely if we need to do things with plots, but make them immutable at first?
  - md5 or other "id" so that we can see if the upstream source has changed.  This is different for googlesheets where the id is properly baked into the sheet

* low level packages:
  - googlesheets
  - excelr
  - these depend on linen, and will have to provide things like ids and filenames to satisfy all the features that linen will do.

* jailbreakr
  - uses output in linen format that is provided by googlesheets or excelr

# Ideas

Can we feed things through openrefine or something?
