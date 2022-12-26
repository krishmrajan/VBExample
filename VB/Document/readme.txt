This Visual Basic program contains examples of how to use the
OLE automation properties and methods for the document and 
folder objects.  It uses most of the methods and properties on 
the Document object, and to a lesser extent, on the Folder object as
well.

Instructions:

1. From Visual Basic, load Document.vbp and run the program 
2. Using the TreeView and ListView controls at the top of the 
   form, select a library and navigate through the folder hierarchy.
   In order to enable the various buttons in the bottom of the form, 
   select either a single document or a single folder. These buttons
   will be enabled based on the capabilities of the selected object.
3. Most of the buttons in the bottom half of the form are self-evident.
   The various 'Show...' buttons will generally bring up forms which
   display the current states or values of the properties for the 
   selected object.  These will not be discussed in any detail here - 
   you can just experiment by clicking on the buttons. 
4. The 'Add' button allows you to enter a new document into the 
   current library.  It will initiate the Add Wizard, which will guide
   you through this process.  You will be asked to select a file from
   your local or network file system, pick a document class, enter 
   document property values, set access rights, then commit the document.
5. The 'Open' button will launch the native application for the selected 
   document. 
6. The 'Check Out' button lets you check out a versioned document from 
   an IDMDS library.  You can optionally choose the directory location
   and local filename for that document once it is checked out.  Once this
   is done, you will observe that the state of the document in the library 
   has changed to 'checked out.'
7. Once a document has been checked out, it can be checked back in again 
   by clicking the 'Check In' button.  You will again have the option of 
   specifying a non-default directory location and filename identifying the
   document to be checked in.  The disabled "full path" text field shows
   where the document was originally placed when it was checked out.  Doing a
   check-in in this way automatically increases the version number of the 
   document. 
8. The 'Send' button will invoke a wizard for creating an E-Mail message 
   with the selected document as an attachment. 

Programming Notes 

The main form contains IDMTreeView and IDMListView controls at the top. 
This application uses them in the conventional way to show the hierarchy of 
libraries and folders in the left pane, with details in the right hand pane.
Single-clicking on an item in the IDMListView pane enables the buttons 
in the form based on returns from the GetState method.  

Most of the button logic is handled in a common way, reflecting the 
uniformity of the underlying object model.  In most cases, a properties 
collection is obtained for the selected object, then the values of 
those properties are displayed in a form.  The mapping between the 
buttons and the associated properties collections are as follows:

    Button Label        Underlying Property Collection
    ------------------  ---------------------------------------------------
    'Versions'          VersionSeries 
    'Permissions'       Permissions
    'Annotations'       Annotations
    'Class'             Document.ClassDescription.ClassPropertyDescriptions
    'Folders Filed in'  FoldersFiledIn

The 'Open' button simply calls the Launch method on the selected document, just 
as the 'Send' button triggers a call on the Send method. The 
'State' button triggers a sequential use of the GetState method using all 
the state enumeration values. 

The Check Out and Check In logic makes use of the Version object, which
is a property of the selected document.  The various text items for 
directory and file names are then passed as parameters to the CheckIn 
and CheckOut methods of the Version object. 

The logic associated with the Properties dialogs provides some examples of how
to work with metadata in the system.  For the selected object, you start by 
getting the property descriptions and property values associated with it.  This
is displayed in the PropertiesForm.  If a single property item is selected
in this listbox, you can then obtain all the state information associated with 
that property - whether it is a required item, has a choice list, is a key, 
etc.  This is accomplished by sequential calls on the GetState method 
of the PropertyDescription object.  Finally, if the property has a choice
list, the 'Choices' form can be launched to display the first page of 
possible values.  This is accomplished by using the Choices property on 
the PropertyDescription object, and by using its NextPage method.  

