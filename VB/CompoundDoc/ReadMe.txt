Welcome to the compound documents sample.

Before you do anything else, open the file called "Tour.doc"

DO NOT RUN THE SAMPLE CODE UNTIL YOU HAVE TAKEN THE TOUR OF COMPOUND DOCUMENTS. THE SAMPLE CODE WILL NOT MAKE ANY SENSE UNTIL YOU HAVE AN UNDERSTANDING OF HOW COMPOUND DOCUMENTS ARE CREATED.

The following areas are covered by this sample:

* Recognizing a compound document using the document properties
* Adding a compound document with the standard user interface
* Adding a compound document without a user interface
* Examining the hierarchy of a compound document
* Opening a compound document with/without a user interface
* Checking in a compound document with/without a user interface
* Checking out a compound document with/without a user interface
* Modifying the properties of an action

NOTE: This sample assumes that you have a document class named "General". If you do not, you will need to edit the file frmMain.frm (in the subroutine mnuAdd_Click).



Instructions:

1. Open the "CompDocSample.vbp" project file and run the program.

2. The Options menu allows you to enable or disable the user interface for compound documents. When you examine the code, you'll see different ways of handling the compound document commands based upon the showUserInterface flag.

3. The status bar at the bottom of the window shows the type of document selected. If the document is a parent in a compound document, the number of children will be displayed in the second status pane (this is the number of direct descendents, not the total number of descendents).

4. The Folder menu allows you to add a compound document to the currently selected folder. If no folder is selected, this menu is grayed out.

5. The Document menu allows you to Display Hierarchy. This shows you the parent-child relationships for the compound document. The code that implements this function is a good example of how to examine the properties of a compound document. 

6. The Document menu also allows you to Open, Check Out and Check In. These functions are implemented with Command, Behavior and Action objects. The code will explain more about these new objects and show you how they are used.


Please review the code for more information.
