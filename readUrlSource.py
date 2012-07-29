#/usr/bin/python2.7

from win32com import client
import urllib, urllib2, os, time

word = client.Dispatch("Word.Application")

# link to the list of files
url1 = 'http://xyz.com/list.php'

# link to delete an entry from the above list
url2 = 'http://xyz.com/del.php'


def printWordDocument(filename):
    """
        Open the document and print it
    """

    word.Documents.Open(filename)
    word.ActiveDocument.PrintOut()
    time.sleep(2)
    word.ActiveDocument.Close()


while True:
    u = urllib
    print "\nURL provided: {0}".format(url1)
    
    # get the document list
    src = u.urlopen(url1)

    # download each document and print them
    for line in src.read().split('\n'):
        print "content of line: " + line
        time.sleep(4)
        if line != '':
            z = ''.join(('http://xyz.com/uploads/',line))
            print z + 'downloading....'
            os.system("c:\Python27\wget.exe {0}".format(z))
            print z + 'downloaded'

            z = ''.join(('c:\\Python27\\',line))
            print "printing " + z + "...."
            
            printWordDocument(z)
            print "deleting " + line
            time.sleep(4)
 
            values = {'link' : line,}
            
            data = urllib.urlencode(values)
            req = urllib2.Request(url2, data)
            response = urllib2.urlopen(req)
        else:
            print "retrying ......"
            time.sleep(3)  

