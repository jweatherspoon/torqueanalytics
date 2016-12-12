import email
import imaplib
import mimetypes
import os
import glob
import zipfile

EMAIL = #Put gmail address here. Ensure access for less secure apps is enabled
PASS = #put password here

def GetMostRecentEmail():
	M = imaplib.IMAP4_SSL('imap.gmail.com')
	#Try to log in
	try:
		M.login(EMAIL, PASS)
	except:
		return
	
	#Get inbox
	rv, data = M.select("INBOX")
	if rv != 'OK':
		return
	
	rv, data = M.search(None, "(UNSEEN)")
	if rv != "OK":
		return
	
	if data[0] != '':
		data = data[0].split()
		for num in data:
			rv, em = M.fetch(num, "(RFC822)")
			if rv != 'OK':
				continue
			msg = email.message_from_string(em[0][1])
			save_attachment(msg)
	else:
		#Get # emails
		rv, data = M.search(None, "ALL")
		if rv != 'OK':
			return

		#data is the number of the last email
		data = data[0].split()[-1:]
		rv, data = M.fetch(data[0], '(RFC822)')
		if rv != 'OK':
			return

		msg = email.message_from_string(data[0][1])
		save_attachment(msg)
		
	M.close()
	M.logout()
	
def save_attachment(msg, download_folder="datafiles/"):
        """
        Given a message, save its attachments to the specified
        download folder (default is datafiles/)

        return: file path to attachment
        """
        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            att_path = os.path.join(download_folder, filename)

            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return att_path
	
def ExtractCSV():
	zips = glob.glob("datafiles/*.zip")
	for f in zips:
		z = zipfile.ZipFile(f)
		z.extractall() #Extract to data/
		os.remove(f)
	#os.system("unzip datafiles/*.zip")
	#os.system("rm -rf datafiles/*.zip")
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	