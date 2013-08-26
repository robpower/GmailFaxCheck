#!/usr/bin/python
__author__="Rob Power"
__date__ ="$26-aug-2013 14.53.58$"
#
#    Copyright (C) Rob Power 2011-2013
#    This file is part of FaxGratis/GmailFaxCheck.
#
#    FaxGratis/GmailFaxCheck is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    FaxGratis/GmailFaxCheck is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with FaxGratis/GmailFaxCheckr.  If not, see <http://www.gnu.org/licenses/>.
#
#------------------------------------------------------------------------------
#	Filename: GmailFaxCheck.py
#   Program Name: GmailFaxCheck
#		  (FaxGratis project)	
#		  https://github.com/robpower/GmailFaxCheck 
#	 Version: 1.2.0
#         Author: Rob Power <dev [at] robpower.info>
#	 Website: http://blog.robpower.info 
#		  https://github.com/robpower 
#		  https://www.gitorious.org/~robpower
#  Last Modified: 26/08/2013
#    Description: This python script checks Gmail IMAP mail server
#                 for the given account for incoming EuteliaVoip Faxes
#                 and outgoing Faxator receipt.
#
#		  It first checks for EuteliaVoip Faxes under label
#		  "$incoming_folder_check": if any Unread mails is detected,
#		  it saves the attached fax PDF under "$archive_dir",
#		  prints a copy and marks email as read.
#
#		  Then checks for Faxator receipts under  label
#		  "$outgoing_folder_check": if any Unread email is detected,
#		  it saves the attached fax receipt PDF  under "$receipt_dir",
#		  prints a copy and marks email as read.
#
#                 Filename is the format: AAAA-MM-DD_HH.MM_SENDER#_originalfilename
#                 Gmail filter must be configured to archive fax mail and move them to
#                 "folder_check" label.
#
#------------------------------------------------------------------------------
#
#
import getopt
import getpass
import os
import sys
import datetime
import imaplib
import email
import email.Errors
import email.Header
import email.Message
import email.Utils
from time import strftime
import shutil
import subprocess
from ConfigParser import SafeConfigParser
#from time import sleep


Usage = """Usage: %s  --user <user> --password <password> --frequency <polling frequency> <imap-server>

	--user		  Provide <user> for authentication on <imap-server>

	--password    Password for the given user

	 Example:  attdownload.py --user username --password password

"""
# Loads settings from "settings.conf" file
settings = SafeConfigParser()
settings.read('settings.conf')
# ARCHIVE
AttachDir = settings.get('archive','AttachDir')		# Attachment Temporary Directory Path
ReceivedArchiveDir = settings.get('archive','ReceivedArchiveDir')	
ReceiptsArchiveDir = settings.get('archive','ReceiptsArchiveDir')	
# GMAIL
incoming_folder_check = settings.get('gmail','incoming_folder_check')
receipts_folder_check = settings.get('gmail','receipts_folder_check') #Faxator Receipts Gmail Label
User = settings.get('gmail','User')			# IMAP4 user
Password = settings.get('gmail','Password')		# User password
# EXTRA
DeleteMessages = settings.get('extra','DeleteMessages')
SaveAttachments = settings.get('extra','SaveAttachments')	# Save all attachments found
Frequency = settings.get('extra','Frequency')			# Mail server polling frequency
exists = settings.get('extra','exists')
name = settings.get('extra','name')
set_read = settings.get('extra','set_read')                    # Put 1 for normal use, 0 for test purpose (does not mark email at end)
DEBUG = settings.get('extra','DEBUG')			# Put 1 for debug output

def usage(reason=''):
	sys.stdout.flush()
	if reason: sys.stderr.write('\t%s\n\n' % reason)
	head, tail = os.path.split(sys.argv[0])
	sys.stderr.write(Usage % tail)
	sys.exit(1)

def args():
	try:
		optlist, args = getopt.getopt(sys.argv[1:], '?',['user=', 'password='])
	except getopt.error, val:
		usage(val)

	global SaveAttachments
	global User
	global Password
	global Frequency

	for opt,val in optlist:
		if opt == '--user':
			User = val
		elif opt == '--password':
			Password = val
		else:
			usage()

	if len(args) != 1:
		usage()

	return args[0]

def write_file(filename, data):
	fullpath = os.path.join(AttachDir, filename)
	fd = open(fullpath, "wb")
	fd.write(data)
	fd.close()

def archive_file(filename, status):
        if status == 'RECEIVED':
            shutil.copy(os.path.join(AttachDir, filename), os.path.join(ReceivedArchiveDir, filename)) # Copies the file to received archive folder
        elif status == 'SENT' or status == 'NOT_SENT' or status == 'CONVERTED' or status == 'Unknown':
            shutil.copy(os.path.join(AttachDir, filename), os.path.join(ReceiptsArchiveDir, filename)) # Copies the file to receipts archive folder

def print_file(filename,dir):
        subprocess.Popen(['lpr', os.path.join(dir, filename)]) # Launches Fax Printing
        os.remove(os.path.join(AttachDir, filename)) # Remove the file from temp directory

def gen_filename(name, mtyp, number, date, status):
        """

        """

	timepart = date.strftime('%Y-%m-%d_%H.%M')
	file = email.Header.decode_header(name)[0][0]
	file = os.path.basename(file)
	path = timepart + '_' + number + '_' + status + '_' + file

	return path

def error(reason):
	sys.stderr.write('%s\n' % reason)
	sys.exit(1)

def walk_parts(msg, number, date, count, msgnum, status):
	for part in msg.walk():

		if part.is_multipart():
			if DEBUG == 1:
				print "Found header part: Ignoring..."
			continue
		dtypes = part.get_params(None, 'Content-Disposition')
		if not dtypes:
			if part.get_content_type() == 'text/plain':
				if DEBUG == 1:
					print "Found plaintext part: Ignoring..."
				continue
			if part.get_content_type() == 'text/html' :
				if DEBUG == 1:
					print "Found HTML part: Ignoring..."
				continue
		else:
			if DEBUG == 1:
				print "Found possible  attachment [Type: " + part.get_content_type() + "]: Processing..."
			attachment,filename = None,None
			for key,val in dtypes:
				key = key.lower()
				if key == 'filename':
					filename = val
					if DEBUG == 1:
						print "[Filename: " + filename + "]"
				if key == 'attachment' or  key == 'inline':
					attachment = 1
					if DEBUG ==1:
						print "[Attach. type: " + key + "]"
				else:
					if DEBUG == 1:
						print "Key: " + key
			if not attachment:
				continue
			filename = gen_filename(filename, part.get_content_type(), number, date, status)

		try:
			data = part.get_payload(decode=1)
		except:
			typ, val = sys.exc_info()[:2]
			warn("Message %s attachment decode error: %s for %s ``%s''"
				% (msgnum, str(val), part.get_content_type(), filename))
			continue

		if not data:
			warn("Could not decode attachment %s for %s"
				% (part.get_content_type(), filename))
			continue

		if type(data) is type(msg):
			count = walk_parts(data, number, date, count, msgnum)
			continue

		if SaveAttachments:
			exists = "0"
			try:
				#Check if its already there
                                if not os.path.isfile(os.path.join(AttachDir, filename)) :
                                    exists = "1"
				if exists == "1":
					write_file(filename, data) # Writes file to temp dir (AttachDir)
                                
				if status == 'RECEIVED':
                                    if not os.path.isfile(os.path.join(ReceivedArchiveDir, filename)):
                                        archive_file(filename, status)
                                    print_file(filename,ReceivedArchiveDir)
	                            print "[" + strftime('%Y-%m-%d %H:%M:%S') +  "]:  Printed Fax  " + filename.split('_')[-1] + "  from  " + number + "  received on  " + date.strftime('%Y/%m/%d - %H:%M:%S')
        	                    print "[" + strftime('%Y-%m-%d %H:%M:%S') +  "]:  Filename: " + filename
                	            print "[" + strftime('%Y-%m-%d %H:%M:%S') +  "]:  ---"
                                elif status == 'SENT' or status == 'NOT_SENT' or status == 'CONVERTED' or status == 'Unknown':
                                    if not os.path.isfile(os.path.join(ReceiptsArchiveDir, filename)):
                                        archive_file(filename, status)
                                    print_file(filename, ReceiptsArchiveDir)
                                    print "[" + strftime('%Y-%m-%d %H:%M:%S') +  "]:  Printed Fax Receipt " + filename.split('_')[-1] + ": fax to  " + number + "  sent on  " + date.strftime('%Y/%m/%d - %H:%M:%S')
                                    print "[" + strftime('%Y-%m-%d %H:%M:%S') +  "]:  Filename: " + filename + '\tResult: ' +status
                                    print "[" + strftime('%Y-%m-%d %H:%M:%S') +  "]:  ---"

			except IOError, val:
				error('Could not create "%s": %s' % (filename, str(val)))

		count += 1

	return count


def process_message(text, msgnum,folder_to_check):

	try:
		msg = email.message_from_string(text)
	except email.Errors.MessageError, val:
		warn("Message %s parse error: %s" % (msgnum, str(val)))
		return text
        date_string ='' + msg["Date"]
        date_string = date_string.split(', ', 1)[1] # Strips out Weekday
        date_string = date_string.split(' +', 1)[0] # Strips out UTC
	date = datetime.datetime.strptime(date_string, '%d %b %Y %H:%M:%S') # Decode datetime from string
	number = msg["Subject"]

	if folder_to_check == incoming_folder_check:
            number = number.split('numero ')[-1]
            status = 'RECEIVED'
	elif folder_to_check == receipts_folder_check:
            if 'OK' in number:
                number = number.split('OK ')[-1]
                status = 'SENT'
            elif 'ERRATA' in number:
                number = number.split('ERRATA ')[-1]
                status = 'NOT_SENT'
            elif 'CONVERSIONE' in number:
                number = number.splt('CONVERSIONE ')[-1]
                status = 'CONVERTED'
            else:
                number = 0000000000
                status = 'Unknown'

	attachments_found = walk_parts(msg, number, date, 0, msgnum, status)
	if attachments_found:
		if DEBUG == 1:
			print "Attachments found: %d" % attachments_found

		return ''
	else:
		if DEBUG == 1:
			print "No attachments found"
		return None


def read_messages(fd):

	data = []; app = data.append

	for line in fd:
		if line[:5] == 'From ' and data:
			yield ''.join(data)
			data[:] = []
		app(line)

	if data:
		yield ''.join(data)


def process_server(host,folder_to_check):

	global DeleteAttachments

	try:
		mbox = imaplib.IMAP4_SSL(host)
	except:
		typ,val = sys.exc_info()[:2]
		error('Could not connect to IMAP server "%s": %s'
				% (host, str(val)))
	
	if DEBUG==1:
		print mbox

	if User or mbox.state != 'AUTH':
		user = User or getpass.getuser()
	if Password == "":
		pasw = getpass.getpass("Please enter password for %s on %s: "
						% (user, host))
	else:
		pasw = Password

	try:
		typ,dat = mbox.login(user, pasw)
	except:
		typ,dat = sys.exc_info()[:2]

	if typ != 'OK':
		error('Could not open INBOX for "%s" on "%s": %s'
			% (user, host, str(dat)))

	if DEBUG == 1:
		print "Selecting Folder " + folder_to_check

	sel_response = mbox.select(folder_to_check)
	#mbox.select(readonly=(DeleteMessages))

	if DEBUG == 1:
		print sel_response

	typ, dat = mbox.search(None, "UNSEEN")

	if DEBUG == 1:
		print typ, dat

	#mbox.create("DownloadedMails")

	#archiveme = []
	for num in dat[0].split():
		typ, dat = mbox.fetch(num, "(BODY.PEEK[])")
		if typ != 'OK':
			error(dat[-1])
		message = dat[0][1]

		if process_message(message, num,folder_to_check) == '':
                        if set_read == 1:
                                mbox.store(num, '+FLAGS', '\\Seen') # Mark Email as Read
	#		archiveme.append(num)
	#if archiveme == []:
	#	print "\n"
	#	print "No mails with attachment found in INBOX"


	#archiveme.sort()
	#for number in archiveme:
	#	mbox.copy(num, 'DownloadedMails')
	#	mbox.store(num, '+FLAGS', '\\Seen') # Mark Email as Read


	#mbox.expunge()
	mbox.close()
	mbox.logout()

process_server('imap.gmail.com',incoming_folder_check)
process_server('imap.gmail.com',receipts_folder_check)
