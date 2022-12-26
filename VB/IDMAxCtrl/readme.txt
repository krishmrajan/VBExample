This Visual Basic program contains examples of how to use two of the
major IDM ActiveX controls: IDMListView and IDMViewer. It also 
uses the ADO mechanism for retrieving some sample documents. 

Instructions:

1.  From Visual Basic, load the IDMAxCtrl.vbp project file and run the 
    program. 
2.  Using the combo box controls at the top of the form, select 
    a library and then a document class identifier. Because this sample
    uses annotations extensively, the library list is currently limited
    to IDMIS libraries.
3.  Click on the 'Find now' button to execute a search for documents
    of the selected class.  The query will return a maximum of 10 
    documents in this case just for use as sample documents. 
4.  The documents matching the search condition will be displayed in 
    IDMListView.  By clicking the right mouse button inside the 
    IDMListView area, you can control the appearance of the displayed 
    items - large icons, small icons, list, or detail.  
5.  If you select a document in the IDMListView, you can again use 
    the right mouse button to bring up a context menu for the document.
    That menu will allow you to view the document, view or modify its
    properties, etc.  Note that this behavior is a standard part of 
    the IDMListView control - the sample application is not involved.
6.  Double-clicking on a document in the IDMListView area causes the 
    sample application to enable the IDMViewer control (in the middle 
    of the form) and load it with the first page of the document.  Again, 
    the Viewer control supports the use of a right mouse button click 
    to bring up a context menu. If the document being viewed is an image,
    this context menu will support operations such as scaling and rotation.
    Although the viewing capabilities are extensive, the sample application
    is not involved in this behavior - it simply "comes for free."
7.  If the 'Show Annotations' checkbox is checked, the sample application
    will get all the annotations for the currently displayed document and page
    and show them in another IDMListView control at the bottom of the 
    form.  It will enable the four buttons to the right of the 
    annotations list to allow creation of new annotations.  The Viewer 
    control will also be set to display annotations.  
8.  Since the lower control is an IDMListView, it supports the right 
    mouse button context menus, this time for looking at properties of 
    the selected annotation.  If you double-click on an annotation item in 
    this control, the Viewer control will display the page which hosts that
    annotation. 
9.  The buttons for adding annotations support four pre-defined 
    annotation types: an "Approved" stamp, a "Reject" stamp, a highlight,
    and a text note.  This is simply an arbitrary set of choices to 
    keep the sample application simple. 
    
Programming Notes

This is a rather small sample application, and all of its functions are
contained in the single main form.  The application demonstrates the 
power of the IDMListView and IDMViewer controls, since most of the 
observed functionality is provided by those controls.  Once a library 
and, optionally, a document class, are selected from the combo boxes, the sample 
application issues a query on the library, using a VB class which 
encapsulates the query details.  This class, called clsSimpleQuery, 
demonstrates how to construct an ADO connection, set properties 
on the connection and the ADO result set, and issue the query. 
At present, the query condition is tied to the 
behavior of an IDMIS, but this could easily be generalized.  

Once the record set is returned from the query operation, the top 
IDMListView control is populated with the results.  This activity is 
also encapsulated in the clsSimpleQuery class, since it is a common 
function to perform.  The internal subroutine 'ShowResults' can be 
reviewed to see the steps involved.  

The entire document viewing behavior is achieved simply through setting
the Document property on the IDMViewer control and setting its 
ShowAnnotations property. 

Once a document page is loaded in the Viewer, the ShowAnnotations function
is called to populate the bottom IDMListView.  The Annotations collection
property is obtained from the Document object and is bound to the 
IDMListView control through the AddItems method.  

Annotation creation is achieved simply through use of the CreateAnnotation
method on the Document object.  Annotation positions are set through 
property values.  
