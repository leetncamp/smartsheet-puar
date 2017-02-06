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
import datetime
import base64

now = "LA.localize(datetime.datetime.now()){0}".format(log_format)


parser = ArgumentParser(description=u"Notify smartsheet owners of non-HP shares")
parser.add_argument("filename", nargs="?", default=default_filename, help=u"File to read.  Must be XLSX format. Set default in confuration.py")
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

log = file("log.txt", "a")

headers_in_email = [u'Name', u'Type', u"Owner", u'Shared To', u'Shared To Permission', "Last Modified (PT)"]
template = Template(open(u"template.html", 'rb').read())
emails_sent = 0
for owner, rows in owners.iteritems():

	for row in rows:
		#normalize timezone to Pacific Time rather than UTC
		row["Last Modified (PT)"] = LA.normalize(UTC.localize(row["Last Modified Date/Time (UTC)"])).strftime(date_format)
		#row["owner_email"] = owner.replace(u"@hp.com", u"")

	firstname = owner.split(".")[0].capitalize()
	html = template.render(**locals())
	To = redirect_emails_to if redirect_emails_to else owner
	msg = Message(To=To, From=From)
	msg.Subject = Subject
	msg.Html = html
	#file("/tmp/delme.html", "wb").write(html)
	msg.customSend("smtp-auth.snl.salk.edu", "nips-assist", base64.b64decode(format))
	emails_sent += 1
	dt = eval(now)
	log.write(u"{0}\tEmailed {1} : {2}".format(dt, owner, html))
	print(u"Sending to {0}".format(owner))
	if emails_sent >= stop_after:
		break

print(u"Sent {0} emails".format(emails_sent))