The HRDemo VB sample application is intended as a demo tool more than as 
a programming guide to the IDM environment.  Written as a skeletal 
human resources application, the sample demonstrates a business-oriented 
interface that uses both types of back-end libraries.  Further, the 
end-user is shielded from the details of those library services, including
the need to log on to them.  The application has been structured to 
allow easy modification or extension to show other types of applications.  
Any dependencies on back-end data structures, such as folders and 
document classes, are contained in a single form that can be modified 
at run-time. 

Basic Concepts:

The basic idea of the sample is to provide services for a human resources
administrator.  The user can examine "new" resumes, search for existing
resumes, or add new ones to the library.  These resume-oriented functions
use the IDMIS library.  You can also examine or update benefits documents
organized in folder structures on the IDMDS library.  Of course, the use
of the IDMIS and IDMDS libraries is arbitrary and invisible to the end-user. 

Setup Instructions: 

The sample uses both IDMIS and IDMIS libraries.  It uses a folder on 
the IDMIS library to provide a set of "new" resumes.  Resume documents
should be scanned into the IDMIS library and filed in the target folder.
The sample also uses a folder hierarchy on the IDMDS library for storing
benefits information.  Suitable documents should be added to the IDMDS
library and filed in the appropriate folders.  

The first time the application is run on a client system, you will be asked
to provide runtime information needed to communicate with the back-end 
libraries: 
	IDMIS Library: the IDMIS library name which is a two-part 
		name comprised of <Library>:<Organization> (e.g. HRLib:FileNET)
	IDMIS User, IDMIS Password: the usercode/password strings needed to 
		log on to the IDMIS library.  Note that a logon will be 
  		done only if the target library is not already logged on; so 
		you can alternatively logon through Explorer or through 
		other applications.  The logon user interface is not displayed.
	Resume folder: the name of the folder on the IDMIS library that
		contains the "new" resumes (e.g. /ResFolder, /New/ResFolder).
		This will be used by the "Find New Resumes" function. 
	Resume DocClass: the name of the document class to be used when 
		querying for resumes or when adding a new resume.  This will 
		be used by the "Find Resume" and "Save New Resume" functions. 
	IDMDS Library: the IDMDS library name, which is a two-part 
		name comprised of <Library>^<Server name> (e.g. QALib^QAMezz) 
	IDMDS User, IDMDS Password: the usercode/password strings needed to 
		log on to the IDMDS library.  Note that a logon will be 
  		done only if the target library is not already logged on; so 
		you can alternatively logon through Explorer or through 
		other applications. The logon user interface is not displayed.
	Benefits Folder: the root folder containing benefits documents
		(e.g. /Benefits, /HRFolder/Benefits).  For example, you might
		construct a folder hierarchy of Benefits/Medical, Benefits/Dental,
		Benefits/Holidays, etc.  This folder hierarchy will be displayed as
		part of the "Benefits" function. 

These settings are stored in the Windows Registry.  The dialog can be accessed 
subsequently by hitting the right mouse button on the main dialog.  New settings 
are used immediately. 

Operating Instructions: 

Once the setup dialog has been completed, the main application form is 
displayed.  There are four large buttons displayed, but only the first two 
are implemented.  Clicking on the "Resumes" button launches a new form 
that supports searching and browsing of resumes, as described below.   

Resume Functions

Clicking on the "Find new" button searches the configured "Resumes" folder 
and populates the IDMListView with documents.  Double-clicking on a document
will cause its contents to be displayed in the IDMTreeView control.  If it
is a scanned image, the annotation buttons are enabled for highlighting and 
creation of notes. The "Request interview" button launches the WorkFlo wizard, 
passing the selected document as an attachment. Note that Ensemble must be 
installed in order for this function to work.  

The "Find Resume" launches a query form to support searching for resumes.
The properties shown in this dialog are those associated with the "resume
document class" you configured in the setup dialog.  If multiple property 
values are specified, they are "ANDed" together to form the final query 
condition.  

The "Add new resume" button can be used to load a local file into the viewer
control.  At this point, the "Save new" button is enabled.  If this button
is clicked, the "Document Add" wizard is launched in order to add the new
document into the IDMIS library.  Note that this file is not currently filed
in the "Resumes" folder, but this can be changed rather easily.  However, in many demo 
situations, you will be quickly adding documents that aren't really 
resumes, and you probably don't want to see them again when you use the 
"Find new" function.  

Benefits Functions

Clicking on the "Benefits" button in the main panel launches a dialog for 
viewing and editing documents.  The TreeView panel is populated with the 
sub-folders contained in the "Benefits" folder you specified in the setup dialog.  
Selecting (opening) a folder in the TreeView will result in a display of its 
contained documents in the ListView panel.  If a document is double-clicked in 
the ListView, it will be displayed in the Viewer control at the right.  

If the selected document can be checked out, the "Update" button will be enabled.
Clicking on this button will check out the document and launch it in 
its native application (e.g. Word, WordPad, etc.)  You can make changes to the 
document, save it to the local file system, then exit the parent application.  At 
this point, the ListView will show the document to be checked out (a red checkmark 
is shown).  Clicking on the "Save" button now will check the document back into the 
system, creating a new version.  The updated document will be refreshed in the Viewer
control, showing the changes you just made.  

Other Functions 
The "401K" and "Stock Purchase" buttons are currently stubbed out.  However, these 
functions can easily be added within the basic structure of the sample application.
In fact, the entire application can be changed fairly easily to mock up an entirely 
different kind of business scenario.  
