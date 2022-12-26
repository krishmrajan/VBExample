This Visual Basic program contains examples of how to use the
OLE automation properties and methods on the Library object.

Instructions:

1. From Visual Basic, load the Library.vbp project and run the program 
2. Select a FileNet library from the IDMListView at the top of the 
   form and log on if necessary. 
3. The form is layed out in a grid of four functional areas.  At 
   the upper right are radio buttons for selecting the type of 
   object you wish to examine.  Clicking on one of these buttons 
   affects the behavior in all other areas of the form.  The default
   selection is 'document.'  These radio buttons are enabled based 
   on the capabilities of the library you selected. 
4. In the lower left and lower center portions of the form, you will 
   see areas which display various class and property related information.
   Clicking on the Show Classes button will retrieve the classes 
   for the type of object you have chosen - document, folder, etc.  
   For example, if the document radio button is set, Show Classes 
   will display all of the document classes configured in the library
   ("document types" for IDMDS).  These class names are displayed 
   in the left-most listbox.  
5. Clicking on a class name in the left-most listbox will populate 
   the center listbox with information about that class. This is 
   property information associated with the class - name, data type, and 
   whether it is a required property. 
6. In the lower-right hand portion of the form, you can retrieve 
   property information on a specific library object, assuming you 
   know its identifier.  For Folder object, the Object ID can be a 
   Folder ID or Name (e.g., 000331451 or /afolder/subfolder).  Again,
   the type of object will be controlled by the radio buttons in the
   top part of the form.  Since this retrieval is associated with
   a specific object, its property values will also be displayed in
   the small listbox in the lower right hand part of the form.
   Of course, there are far more sophisticated ways of finding objects
   in the system, but these rely on methods which are not part of the
   Library object.

Programming Notes

As advertised by its name, this application focuses almost exclusively 
on the properties and methods of the Library object.  All of the 
logic is contained in the single, main form.  The top IDMListView control is 
populated with the Libraries collection from the Neighborhood object.  
In fact, this is probably the most economical way to present a user 
with a list of libraries - one line of code. Once a library object is 
selected, it becomes the focal point for most of the other behavior in
the application. 

The radio buttons in the form are enabled based on calls to the 
Supports method on the Library object.  When the Show Classes button 
is clicked, the FilterClassDescriptions method is called on the 
Library object to populate the leftmost listbox.  This call is made 
using the object type required by settings of the radio buttons. 

When a particular class item is selected in the leftmost listbox, a
call is made on FilterPropertyDescriptions, passing the selected class
name as a parameter.  The resulting collection is used to populate 
the central listbox.  

The behavior in the lower right hand portion of the form is controlled 
by capabilities of the GetObject method on the Library object.  The
logic in the form assumes that the identifier entered by the user is 
the ID or Name.  For IDMDS libraries, the object ID is passed as a string.
If a version number is specified, it is appended to the base identifier
to specify retrieval of a specific version of the object. 
