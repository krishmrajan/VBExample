This Visual Basic program is a simple example of a document entry program 
which does not rely on the "Add Wizard."  Relying primarily on the 
Document object, it shows techniques for obtaining document class 
information, setting property values, and creating a new document 
in a library.  It also uses the IDMViewer control for viewing 
local files. 

This sample was originally intended as a small demo application. The 
basic idea was to provide a mechanism for a user to browse a set of 
local files, decide which ones should be added ("committed") to the 
IDM library, and then provide the property information necessary to 
save them.  

The buttons on the various forms have been configured to show icons 
for the functions they represent.  In the description below, the 
buttons are identified by their "tooltip" text.

Instructions:

1. Run Visual Basic, load the Document.vbp project, then run the program 
2. You will be prompted with a standard Windows file open dialog box. 
   Select a file or a set of files which you want to browse and 
   optionally add to an IDM library. 
3. Using the VCR controls for 'up' and 'down', you can browse 
   through the list of documents, examining them in the Viewer 
   control.  If the documents are images, you can use the image 
   rotation and scaling buttons below the Viewer pane.  Alternatively, 
   you can use the right mouse button inside the Viewer pane to 
   access other functions. Note that the Viewer control will operate 
   on non-image documents using the format filters included with 
   the control. 
4. In order to add the current document to the library, click on the 
   'commit' button. This will bring up another form which asks you to 
   select a library, a document class and a folder this document to be
   filed.  The folder name can be empty or like /aaa, /aaa/bbb.
   Once you have done this, you will see that the property grid is filled
   with the property names and data types for that document class.  
   Properties marked with a red checkmark are required.  You can enter 
   property values into the grid.  Clicking on the 'Check' button confirms
   that you wish to commit this document.  Clicking on the 'Exit' button
   cancels the commit operation for this document. 
5. After you have entered properties for a document, the status bitmap
   at the top of the form will change to a checkmark.  If you click 
   the commit button again on such a document, it will be dropped from
   the commit list and the bitmap will change to an 'X'. 
6. Clicking on the 'Restart' button at any time will cancel any pending
   operations and will return the program to the initial File Open dialog.
7. Clicking on the 'Exit' button will end program execution without 
   saving any documents to the library. 
8. Clicking on the 'Finish' button will process the list of documents 
   marked for committal and will add them to the selected library. A 
   progress dialog box will be displayed until the entire operation is 
   completed. 
