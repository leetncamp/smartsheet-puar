#!/usr/bin/env python

from excel2dict import excel2dict


from argparse import ArgumentParser
parser = ArgumentParser(description=u"Notify non-HP smartsheet owners")
parser.add_argument("filename", nargs="?", default="Access.xlsx", help=u"File to read.  Must be XLSX format.")
ns = parser.parse_args()

data = excel2dict(ns.filename)
print data