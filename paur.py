#!/usr/bin/env python

from excel2dict import excel2dict
from pdb import set_trace as debug
import traceback
from argparse import ArgumentParser
import re
from jinja2 import Template
import smtplib
import pytz
import os, sys
UTC = pytz.timezone("UTC")
LA = pytz.timezone("America/Los_Angeles")
from snlmailer import Message
from configuration import *  #supplies From and Subject for the email
from collections import OrderedDict


parser = ArgumentParser(description=u"Notify non-HP smartsheet owners")
parser.add_argument("filename", nargs="?", default="Access.xlsx", help=u"File to read.  Must be XLSX format.")
ns = parser.parse_args()

data = excel2dict(ns.filename)

hpRE = re.compile("@hp.com", re.I)

owners = {}
count = 0
for row in data:
	count += 1
	#print(u"processing row {0}".format(count))
	shared_to = row[u"Shared To"]

	if shared_to and not hpRE.search(shared_to):
		owner = row[u"Owner"]
		owners[owner] = owners.get(owner, []) + [row]
		
	else:
		continue #Skip this row if the Shared-to is an hp address

for owner in owners:
	print(owner)

raw_input(u"Sending {0} emails. Press ^C to cancel. Any key to continue...".format(len(owners)))

headers_in_email = [u'Name', u'Type', u'Shared To', u'Shared To Permission', 'Last Modified Date/Time (UTC)']
template = Template(open(u"template.html", 'rb').read())
for owner, rows in owners.iteritems():
	firstname = owner.split(".")[0].capitalize()
	html = template.render(**locals())
	msg = Message(To="lee@snl.salk.edu", From=From)
	msg.Subject = Subject
	msg.Html = html
	file("/tmp/delme.html", "wb").write(html)
	os.system('open /tmp/delme.html')
	msg.snlSend()
	print(u"Sending to {0}".format(owner))
	sys.exit()