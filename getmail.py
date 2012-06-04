import poplib
import email
import os
import time
import base64
import quopri
import StringIO
import rfc822
import getpass

class GmailTest(object):
    def __init__(self):
        self.savedir="c:\\inbox\\"
        self.mailsavedir="c:\\inbox\\"

    def test_save_attach(self):
        self.connection = poplib.POP3_SSL('pop.yandex.ru', 995)
        self.connection.set_debuglevel(1)
        self.connection.user("begininva")
        self.connection.pass_("XXXX")
        self.connection.pass_(pas)
        emails, total_bytes = self.connection.stat()
        print("{0} emails in the inbox, {1} bytes total".format(emails, total_bytes))
        # return in format: (response, ['mesg_num octets', ...], octets)
        msg_list = self.connection.list()
		
        print(msg_list)

        # messages processing
        for i in range(emails):
            # return in format: (response, ['line', ...], octets)
            response = self.connection.retr(i+1)
            raw_message = response[1]
            text1 = '\n'.join(response[1])
            str_message = email.message_from_string("\n".join(raw_message))
            mesg1 = StringIO.StringIO(text1) 
            msg1 = rfc822.Message(mesg1)
            name1, email1 = msg1.getaddr("From")
            d = msg1.getdate('Date')
            emailfile = str(self.savedir +  email1 +" "+ time.strftime("%d_%b_%Y %H-%M-%S",d)  + ".eml")
            print(emailfile)
            file = open(os.path.join(self.mailsavedir, emailfile), 'wb')
            file.write(text1)
            file.close()

            #str_message = email.message_from_bytes(b'\n'.join(raw_message))
            self.connection.dele(i+1)
            #time.sleep(1)
            # save attach
            for part in str_message.walk():
                print(part.get_content_type())

                if part.get_content_maintype() == 'multipart':
                    continue
					
                if part.get('Content-Disposition') is None:
                    print("no content dispo")
                    continue
					
                filename = part.get_filename()
                if not(filename): filename = "test.txt"

                print(filename.split("?"))
                if filename.split("?")[0] in "=" : 
                    if not filename.split("?")[2] == "Q" :
                        filename = unicode( base64.b64decode(filename.split("?")[3]), 'koi8-r' ).encode('cp1251')
                    else:
                        filename = unicode( quopri.decodestring(filename.split("?")[3]), 'koi8-r' ).encode('cp1251')
                
                print(filename)
                #print( unicode( base64.b64decode(filename.split("?")[3]), 'koi8-r' ).encode('cp866'))
                
                fp = open(os.path.join(self.savedir, filename), 'wb')
                fp.write(part.get_payload(decode=1))
                fp.close

        self.connection.quit()
		 
        #I  exit here instead of pop3lib quit to make sure the message doesn't get removed in gmail
        import sys
        sys.exit(0)

d=GmailTest()
d.test_save_attach()