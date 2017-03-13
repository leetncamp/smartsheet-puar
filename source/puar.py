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
from configuration import *  #pulls in variables defined in configuration.py

from collections import OrderedDict
import datetime
import base64
from nameparser import HumanName


stop_after       = locals().get("stop_after", 0) #stop_after defaults to 0 if not present in configuration.py
default_filename = locals().get("default_filename", "UserAccessReport.xlsx")
smtp_server      = locals().get("smtp_server", "smtp3.hp.com")
log_format       = locals().get("log_format", ".isoformat()")
date_format      = locals().get("date_format", "%Y-%m-%d %H:%M")
format           = locals().get("format", 'WVpFLWFkUi02SlEtcm80')

now = "LA.localize(datetime.datetime.now()){0}".format(log_format)

parser = ArgumentParser(description=u"Notify smartsheet owners of non-HP shares")
parser.add_argument("filename", nargs="?", default=default_filename, help=u"File to read.  Must be XLSX format. Set default in configuration.py")
parser.add_argument("--go", action="store_true", help="with --go, the program does not wait for confirmation. It sends email. Useful for automated running. ")
ns = parser.parse_args()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

data = excel2dict(os.path.join(BASE_DIR, ns.filename))

hpRE = re.compile("@hp.com", re.I)

#png_data = base64.b64encode(open("image001.png", "rb").read())

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
if redirect_emails_to:
	print(u"\nRedirecting emails to {0}!".format(redirect_emails_to))

if stop_after:
	print(u"\nStopping after {0} emails".format(stop_after))

if not ns.go:
	raw_input(u"\nFound {0} users needing notification. Press ^C to cancel. Any key to continue...".format(len(owners)))

log = file("log.txt", "a")

headers_in_email = [u'Sheet Name', u'Shared To Permission', u'Shared To']
template = Template(open(u"template.html", 'rb').read())
emails_sent = 0

for owner, rows in owners.iteritems():

	for row in rows:
		#normalize timezone to Pacific Time rather than UTC
		row["Last Modified (PT)"] = LA.normalize(UTC.localize(row["Last Modified Date/Time (UTC)"])).strftime(date_format)
		#row["owner_email"] = owner.replace(u"@hp.com", u"")
		row["Sheet Name"] = row[u"Name"]


	nameStr = u" ".join([item.capitalize() for item in owner.split(u"@")[0].split(".")])
	name = HumanName(nameStr)
	firstname = name.first
	lastname = name.last
	html = template.render(**locals())
	To = redirect_emails_to if redirect_emails_to else owner
	msg = Message(To=To, From=From)
	msg.Subject = Subject
	msg.Html = html
	addlHeaders = [["Precedence","bulk"], ['Disposition-Nofication-To', From]]

	#try:
	#	file("/tmp/delme.html", "wb").write(html)
	#	os.system("open /tmp/delme.html")
	#except:
	#	pass
	#debug()
	#msg.customSend("smtp-auth.snl.salk.edu", "nips-assist", base64.b64decode(format))

	errors = msg.customSend(smtp_server)
	if errors:
		print(errors)
	emails_sent += 1
	dt = eval(now)
	log.write(u"{0}\tEmailed {1}\n".format(dt, msg.To))
	print(u"Sending to {0}".format(owner))
	if stop_after and emails_sent >= stop_after:
		print(u"Stopping early at {0} emails. See configuration.py stop_after".format(stop_after))
		break

print(u"Sent {0} emails".format(emails_sent))
if not ns.go:
	raw_input("Press any key to exit.")