This Visual Basic program contains examples of how to use the
OLE automation properties and methods on the CustomObject object.

Custom objects are supported only by IDMDS libraries at the 
present time.  These objects allow IDMDS to act as a persistent
store for data created and used by third party applications.  They 
are treated my IDMDS as opaque objects - IDMDS doesn't "know" 
what they represent.  They are uniquely identified by a three-tuple: 
ISV Id, Object Type, and Key.  Like all other objects in the library, 
they have properties associated with them that can be stored and 
retrieved, independent from the object data stream itself. 

This sample application is very simple.  It basically allows you to
search for custom objects based on the three-part identifier, examine
their properties, and create new custom objects.  

Instructions:

1. From Visual Basic, load CustObj.vbp and run the program. 
2. You will be prompted with a combo box of IDMDS libraries. 
   Choose one and log on if necessary. 
3  The main form will be initialized with a reserved ISV ID and 
   default values for object type and key.  You can click on the 
   'Get' button to find a custom object whose identifier matches the
   search parameters.  Note that there can be at most one.  If a 
   match is found, the properties listbox is populated. 
4. Clicking on one of the properties in the listbox will result in 
   a display of its current value in the value textbox.  You can also
   obtain the value of an extended property by entering the name of 
   that property and clicking 'Get'.  
5. In order to add a new custom object, you must specify an ISV 
   identifier different from the default.  If you do not, you will 
   get an error message.  However, you can use any string you want for 
   an ISV id - there is nothing special that needs to be configured 
   in the IDMDS server.  If property values have been changed on
   a custom object, clicking on 'Save' will store the new values in 
   IDMDS. 

Programming Notes:

Everything other than the basic library logon is handled in the main
form (Form1).  Retrieving a custom object based on its unique identifier
is accomplished by using the GetObject method on the Library object.  Creating
a new custom object is a two-step operation.  You first do a CreateObject
call on the Library object, specifying the ISV identifier.  You then 
get Property objects for ObjType and ObjKey from the new CustomObject, and
set their Value properties as desired.  The 'Add' function automatically
does a call on the Save method of CustomObject.

The tabbed dialog in the center of the form is used to display properties
and permissions on the current CustomObject.  Populating these fields 
is accomplished by use of the GetState method and the Permissions property. 
At present, you can't modify any of the property values in these dialogs.



