XLS
=======================

[![CircleCI](https://circleci.com/gh/lunny/xls.svg?style=shield)](https://circleci.com/gh/lunny/xls)  [![codecov](https://codecov.io/gh/lunny/xls/branch/master/graph/badge.svg)](https://codecov.io/gh/lunny/xls)
[![](https://goreportcard.com/badge/github.com/rinor/xls)](https://goreportcard.com/report/github.com/rinor/xls)
[![GoDoc](https://godoc.org/github.com/rinor/xls?status.svg)](https://godoc.org/github.com/rinor/xls)

# Installation

    go get github.com/rinor/xls

# Reference

* [http://www.ozgrid.com/Excel/CustomFormats.htm](http://www.ozgrid.com/Excel/CustomFormats.htm)
* [http://www.openoffice.org/sc/excelfileformat.pdf](http://www.openoffice.org/sc/excelfileformat.pdf)

# Contributors

Pure Golang xls library writen by [Rongshu Tech(chinese)](http://www.rongshu.tech).
Add Formula & Format support by [chen.s.g] (http://www.imohe.com)

Thanks for contributions from Tamás Gulácsi, sergeilem.

**English User please mailto** [Liu Ming](mailto:liuming@rongshu.tech)

This is a xls library writen in pure Golang. Almostly it is translated from the libxls library in c.

The master brunch has just the reading function without the format.

***new_formater** branch is for better format for date and number ,but just under test, you can try it in development environment. If you have some problem about the output format, tell me the problem, I will try to fix it.*

# Basic Usage

* Use **Open** function for open file
* Use **OpenWithCloser** function for open file and use the return value closer for close file
* Use **OpenReader** function for open xls from a reader, you should close related file in your own code

* Follow the example in GODOC