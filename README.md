# jailbreakr

[![Linux Build Status](https://travis-ci.org/jennybc/jailbreakr.svg?branch=master)](https://travis-ci.org/jennybc/jailbreakr)
[![Windows Build status](https://ci.appveyor.com/api/projects/status/github/jennybc/jailbreakr?svg=true)](https://ci.appveyor.com/project/jennybc/jailbreakr)

Data Liberator.  To extract tabular data people put in nontabular structures in a program designed to hold tables.

![](http://i.giphy.com/SEp6Zq6ZkzUNW.gif)

## Installation

Requires the development version of xml2 (for `xml_find_lgl`)

```r
devtools::install_github("hadley/xml2")
devtools::install_github("jennybc/jailbreakr")
```

## Goals

There are two large excel spreadsheet corpuses; it would be nice to use these to get a feel for what fraction of spreadsheets we can handle or the range of non-table-like data out there.

![the things people do to data](http://replygif.net/i/514.gif)

The first is the [EUSES corpus](http://openscience.us/repo/spreadsheet/euses.html) of 4,447 spreadsheets (16,853 worksheets).  This is all xls files (rather than xlsx) and therefore need either an [xls -> xlsx conversion](http://bit.ly/1P2rMGr) or support in jailbreakr for xls files.

The second, larger, one is the [Enron corpus](http://www.felienne.com/archives/3634) of 15,770 spreadsheets (79,983)
