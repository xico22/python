# Import system modules
import win32com.client
import pywintypes
import getpass

# Get credentials
mailServer = 'notes_server'
mailPath = r'mail_path'
mailPassword = getpass.getpass()
# Connect
notesSession = win32com.client.Dispatch('Lotus.NotesSession')
try:
    notesSession.Initialize(mailPassword)
    notesDatabase = notesSession.GetDatabase(mailServer, mailPath, False)
except pywintypes.com_error:
    raise Exception('Cannot access mail using %s on %s' % (mailPath, mailServer))

def makeDocumentGenerator(folderName):
    # Get folder
    folder = notesDatabase.GetView(folderName)
    if not folder:
        raise Exception('Folder "%s" not found' % folderName)
    # Get the first document
    document = folder.GetFirstDocument()
    # If the document exists,
    while document:
        # Yield it
        yield document
        # Get the next document
        document = folder.GetNextDocument(document)

# Get a list of folders
def listFolder():
    for view in notesDatabase.Views:
        print view.Name

#Build the view and get first document in view
#for document in makeDocumentGenerator('($Inbox)'):
#    print document.GetItemValue('Subject')[0].strip()


print ', '.join(i for i in dir(notesDatabase) if not i.startswith('__'))
