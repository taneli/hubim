import os
import imaplib
import getpass
import email, email.Utils
#import datetime
import socket
import time
import xml.etree.cElementTree as ET

CONFIG = {
	'IMAPSERVER': "imap." + ".".join(socket.getfqdn().split(".")[1:]),
	'IMAPUSERNAME': os.getlogin(),
	'FOLDERNAME': "Inbox",
	'OUTPUTFILE': "imap.xml",
	'DAYS': "14",
	'DEBUG': 0
	}

def readconfig(fname):
	for line in file(fname, "r").read().splitlines():
		curline = line.strip()
		if curline.startswith("#"):
			continue
		k, v = map(lambda x: x.strip(), curline.split("="))
		if CONFIG['DEBUG']:
			print "DBG: Adding key %s with value %s" % (k, v)
		CONFIG[k] = v

def sort(L):
	L.sort()
	return L

def gethdrs(fieldspec):
	"""Get IMAP headers according to fieldspec, which must follow
	RFC 3501 SEARCH specification."""
	retval = list()
	mbox = imaplib.IMAP4(CONFIG['IMAPSERVER'])
	mbox.login(CONFIG['IMAPUSERNAME'], CONFIG['IMAPPASSWORD'])
	mbox.select(CONFIG['FOLDERNAME'], True)
	typ, nums = mbox.search(None, fieldspec)
	print "INFO: fetching %d headers." % len(nums[0].split())
	for num in nums[0].split():
		typ, data = mbox.fetch(num, "(UID RFC822.HEADER)")
		uid = data[0][0].split()[data[0][0].split().index("(UID")+1]
		retval.append((uid, data[0][1]))
	mbox.close()
	mbox.logout()
	return retval

def generate_feed(feed, hdrs):
	unordered = dict()
	for hdr in hdrs:
		phdr = email.message_from_string(hdr[1])
		elem = ET.Element("entry")
		ET.SubElement(elem, "title").text = phdr['Subject']
		
	# <link href="imap://server/INBOX;UID" rel="alternate"/>
		link = ET.SubElement(elem, "link")
		link.attrib['rel'] = "alternate"
		link.attrib['href'] = "imap://%s/%s;UID=%s" % (CONFIG['IMAPSERVER'], CONFIG['FOLDERNAME'], hdr[0])

	# <id>UID</ID>
		ET.SubElement(elem, "id").text = phdr['Message-ID'].strip("<>")

	# <updated>YYYY-MM-DDTHH:MM:SSZ</updated>
		date = email.Utils.parsedate(phdr['Date'])
		ET.SubElement(elem, "updated").text = time.strftime("%Y-%m-%dT%H:%M:%SZ", date)

	# <author><name>name</name><email>email</email></author>
		author = ET.SubElement(elem, "author")
		name, addr = email.Utils.getaddresses([phdr['From']])[0]
		ET.SubElement(author, "name").text = name
		ET.SubElement(author, "email").text = addr
		unordered[long(time.strftime("%s", date))] = elem
	ordered = sort(unordered.keys())
	for key in ordered:
		feed.append(unordered[key])
	return feed

def do_work():
	datespec = time.strftime("%d-%b-%Y", time.gmtime(time.mktime(time.gmtime()) - long(CONFIG['DAYS']) * 86400))
	feed = ET.Element("feed")
	feed.attrib['version'] = "0.3"
	feed.attrib['xmlns'] = "http://purl.org/atom/ns#"
	ET.SubElement(feed, "title").text = "IMAP %s@%s" % (CONFIG['IMAPUSERNAME'], CONFIG['IMAPSERVER'])
	ET.SubElement(feed, "updated").text = "%s" % (time.strftime("%Y-%m-%dT%H:%M:%SZ"))
	link = ET.SubElement(feed, "link")
	link.attrib['rel'] = "alternate"
	link.attrib['href'] = "imap://%s/%s" % (CONFIG['IMAPSERVER'], CONFIG['FOLDERNAME'])
	generate_feed(feed, gethdrs("(SINCE %s)" % datespec))
	file(CONFIG['OUTPUTFILE'], "w").write(ET.tostring(feed))

if __name__ == "__main__":
	import sys
	if len(sys.argv) != 1:
		print "Usage: imap2atom.py"
		sys.exit()
	RCFILE = os.path.join(os.environ['HOME'], ".imap2atomrc")
	if os.path.exists(RCFILE):
		readconfig(RCFILE)
	if CONFIG['DEBUG']:
		print "Resources: %s" % str(CONFIG)
	CONFIG['IMAPPASSWORD'] = getpass.getpass()
	do_work()
