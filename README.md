# ht2excel

A very simple command-line tool for converting HTML tables from a file into Excel Sheets.
Each table is it's own Excel Sheet...

## Usage

```sh
$ ht2excel -f testdata/test.html -o data/output.xlsx
```

## Building

You will need Go 1.19+ to build it

```sh
$ git clone https://github.com/zikani03/ht2excel

$ cd ht2excel

$ go build 

```
---

Copyright (c) Zikani Nyirenda Mwase