This Visual Basic program contains examples of how to use the
OLE automation properties and methods on the Annotation object.

Instructions:

1. Open Anno.vbp and execute the program 
2. Click on the 'Select Document' button.  This will bring up 
    a common dialog for opening and browsing FileNet libraries. 
    Since annotations are not supported on IDMDS, you need to
    choose an IDMIS library. 
3. Navigate through the open dialog to select a single document. 
    You can do a 'find' on a specific document id if you know it, 
    or you can navigate through the folder hierarchy. For now, the
    advanced 'find' capability has been disabled, but this is 
    temporary. 
4. If the document you open has annotations already, they will be 
    displayed in the listbox.  If the document currently has no annotations 
    you can use the 'Create annotation' button to make one.
5. If you choose to create a new annotation, you will be asked to 
    select its type (text, highlight, arrow, etc.)  You can  
    specify its other properties if you wish.  The annotation will 
    be associated with the current page of the document.  The 
    'next' and 'previous' page controls at the top of the listbox can 
    be used to control which page will host the annotation.   
6. If an annotation is selected in the listbox, you can use the 
    other buttons at the right of the form to modify its security, 
    delete it, or view its properties.  The 'View Properties' button 
    will display a simple listbox showing property values.  The 
    'View Properties Dialog' button will initiate the standard 
    annotation properties dialog, which also allows modification of 
    properties.  

Remember that you are actually manipulating "real" annotations on 
these document pages, even though you aren't seeing the actual pages
of the document.  After running this program and adding or modifying 
annotations, you can see the results of your work using the IDMViewer 
application or one of the other sample programs which uses the Viewer 
control. 

Programming Notes:
This sample deals primarily with the Document and Annotation objects 
and can give you a better understanding of the relationships between
them.  Remember that an annotation is always bound to exactly one 
page of one document, although it may have security that is different
from that of the underlying document. 

Form 1

Logic behind the 'Select Document' button uses the CommonDialog object 
to deal with opening a library, navigating folders, and selecting a 
document.  After this is done, global variables are set for the Document
and Library objects.  Population of the annotation list is accomplished 
through use of the GetPageAnnotations method on the Document object. 

If the user asks to see annotation properties, this can be accomplished
with one line of code using a call on the ShowPropertiesDialog method
on the Annotation object.  Alternatively, they can be displayed in a 
ListView using Form 4. 

Form 2

This form handles the creation of an annotation.  The user must first 
choose the style of annotation, based on the styles supported in the 
library.  This enumeration is accomplished with the a FilterClassDescriptions
call on the Library object, specifying annotation types as the filter 
parameter.  Actual creation of the annotation is again a one line 
operation - just a call on the CreateAnnotation method of the Document
object.  This again emphasizes that annotations are always bound to 
documents.  Modification of properties for the new annotation are 
handled by a call to ShowPropertiesDialog on the Annotation object. 

Form 3 

This form handles modification of security properties.  For performance
reasons, the main form obtains the ObjectSets for groups and users 
associated with the library and stores them globally.  These are 
subsequenly used in Form 3 to populate the lists of available users and groups.
Modification to these properties on the Annotation are accomplished by
setting values ofthe Permissions property.

Form 4

Form 4 is used to display annotation properties in a simple MS ListView, 
as opposed to using the ShowPropertiesDialog.  This demonstrates use of 
the Properties collection on the Annotation object.  
  
Module 1

This module is used to house the error handling logic in the application.
The ShowError function uses the ErrorManager object to display the 
stack of Error objects after an exception has occurred, assuming that 
the error was generated in a FileNet subsystem.

