This Visual Basic sample illustrates the use of three IDM ActiveX controls: 
the IDMTreeView, the IDMListView, and the IDMViewer.  Most of the 
attention in this readme file will be focused on the first two 
controls, since the Viewer control is used heavily in other samples. 

Instructions: 
1.  From Visual Basic, load the "MiniExplorer.vbp" project file, and run the 
    program. 
2.  Select an IDM library from the initial dialog box, and log on 
    if necessary.  The sample will work with either an IDMIS or 
    IDMDS library. 
3.  The main form is divided into three panel areas: the IDMTreeView 
    and IDMListView panels appear in the top half of the form, while
    the IDMViewer occupies the bottom half.  The IDMTreeView panel is 
    populated initially with all top level folders residing in all 
    of the libraries you are logged onto.  Note that this panel may show 
    more than one library if you are logged on to multiple libraries 
    through other applications or through the Windows Explorer.  
4.  You can navigate through the IDMTreeView to see the folder hierarchy. 
    When a folder or library is selected in the IDMTreeView, its contents
    will be displayed in the IDMListView panel.  The IDMListView supports
    the standard right mouse button context menu for any object that 
    is selected in that panel.  This lets you examine the object's properties, 
    check it out, launch the IDMViewer application on it,etc.  Double-clicking
    on the selected item will initiate a display of the document in the 
    IDMViewer control at the bottom of the form.  Here again, the standard
    right mouse button behavior in the IDMViewer control is supported.  
5.  One of the objectives of this sample is to illustrate support for 
    drag and drop operations in the IDM controls.  The following types of 
    operations are supported: 
    - Dragging a folder from one location in the IDMTreeView to another 
    - Dragging a folder from the IDMTreeView to the IDMListView panel
    - Dragging a document or folder from the IDMListView to a folder in the 
      IDMTreeView 
    - Dragging a document from the IDMListView to a folder in the 
      same IDMListView panel 
    - Dragging a document from the Windows Explorer to either the 
      IDMListView or the IDMTreeView 
6.  The sample also supports clipboard copy and paste operations using
    the right context menu in the IDMTreeView panel.  In addition, folder 
    operations such as 'delete' and 'new' are supported through a combination 
    of main menu pulldowns and right mouse button context menus.  
7.  'Delete' operations are constrained to prevent the actual deletion of 
    documents.  Folders can be deleted, and embedded documents will be unfiled, 
    but the documents will not actually be deleted from the library.  Note
    that this is a choice made in the sample application and is not a built-in
    restriction of the Panagon system.  A number of main menu options are 
    "stubbed out", since implementing them would have little value as an 
    illustration of Panagon programming capabilities.    
8.  The sample application also makes use of other Visual Basic capabilities
    that are relevant to real-world applications; including splitter bars,
    externally stored (and translatable) strings, a toolbar, and handling of
    form resize events.  

Programming Notes

Program execution begins in the 'Main' subroutine of 'PublicModule.'  Initialization
code here deals primarily with restoring application settings from the registry.
This module also contains a number of other utility functions for populating the 
IDMListView and IDMTreeView controls, displaying error messages, terminating the
application, and recursively unfiling documents from the folder hierarchy.  It 
also contains the subroutine for loading form resource strings from the ".res" file.
This routine is called from the 'Load' function of each of the forms.  The Panagon-specific 
initialization calls an internal logon routine which lets the user select and 
logon to a library.  For demo purposes, the sample app supports a mode of 
operation in which logon is done automatically using pre-defined usercode and 
password strings.  The dialog box for enabling this feature, called frmOptions, can 
be accessed through a menu option under the 'View' category.  Settings for these 
default logon strings are contained in the resource (.res) file, which can be edited 
using any of the standard resource compiler tools (DevStudio, VC++, etc.)  In a 
departure from the other sample applications, the user is not required to logon to an 
IDM library from within this sample.  Thus, if a user cancels the logon sequence for
any reason, the application will proceed.  

Once the logon sequence is completed, execution continues in frmMain.  At this point,
the IDMTreeView is populated with all IDM libraries and root folders for which the 
desktop is logged on.  Thus, the application can be used to demonstrate a seamless
view of both types of libraries.  This feature also explains why a logon is not forced -
the user may, in fact, have logged on to libraries through other applications or 
through the Explorer shell.  

As you might expect, many user interface capabilities for the IDMTreeView and 
IDMListView controls come "free of charge", so the sample application does not need
to add code for supporting drag and drop or clipboard operations within a control.  The
sample does show how to hook drag events for IDMListView through the subroutines 
"ilvIDMListView_DoBackgroundDragDrop" and "DoBackgroundDragOver".  The latter routine
simply gives the user a visual feedback about whether the dragged object can be 
dropped on the IDMListView panel.  The first subroutine handles the case where a 
folder, document, or string object is being dragged onto the IDMListView from some 
other control area.  For example, dragging a folder from the IDMTreeView and dropping 
it on the IDMListView panel must be handled here.  In general, handling of these types
of actions requires some knowledge about the semantics of the operation.  For example, 
dragging a folder in this way may produce the same result as dragging and dropping 
the folder within the confines of the IDMTreeView control (i.e a "folder copy"). 
 
The sample also demonstrates how events in the IDMTreeView control can be "hooked" 
using the "BeforeInvokeCommand" and "DoInvokeCommand" events.  In this case, the 
logic is actually working around some difficulties with repainting the tree control.
However, the example illustrates how to hook these events and implement application-
specific behavior.  

The sample application also provides support for events such as item deletion, item 
creation, and many of the cut/copy/paste clipboard operations.  In general, these
subroutines must be added in order to provide the semantic support implied by the 
actions of the user interface control.  One particular example is the interpretation
applied to a folder delete.  This sample has chosen to interpret such an event as a 
"unfile documents/delete folder" operation.  This behavior is implemented in 
subroutines that respond to the "ItemDelete" events in both the IDMListView and 
IDMTreeView controls.  Similarly, support is added for the single and double-click 
events of the IDMListView control, whereby single-click updates certain UI controls
and double-click loads a document into the IDMViewer control.  Object renaming is 
accomplished by catching and handling the "EditLabel" family of events. 

In general, the programmer must consider two aspects to these control events: the 
desired behavior in the user interface and the behavior desired in the IDM library. 
The general approach taken in the sample app is to invoke the correct behavior on 
the IDM object (e.g. unfile a document, move a folder, delete a folder), then 
refresh the user interface controls to reflect the new state of the IDM world.

