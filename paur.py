#!/usr/bin/env python

from openpyxl import load_workbook


from argparse import ArgumentParser
parser = ArgumentParser(description=u"Notify non-HP smartsheet owners")
parser.add_argument("filename", nargs="?", default="Access.xlsx", help=u"File to read.  Must be XLSX format.")
ns = parser.parse_args()

wb = load_workbook(ns.filename)



