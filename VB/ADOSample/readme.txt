This is a low-level sample program which demonstrates how to use ADO objects
to query documents on IDMIS and IDMDS servers.  This sample is
provided to show programming techniques using ADO objects and IDM
objects. 

Instructions:

(1) Start Visual Basic, open the ado.vbp project file and run the application.

(2) Select the appropriate library from the combo box and log on if necessary. 

(3) The 'NewCommand' menu will be enabled.  Click it to bring up
    the 'CreateCommand' form.  You can also click the 'Properties' item
    in the 'Connections' pulldown to examine the global properties of 
    the ADO connection.

(5) The layout of the Command form is different for IDMIS and IDMDS.
    For either back-end, you can type in a SQL command or click one of
    the small buttons under the edit box to copy sample SQL commands
    into the SQL edit box.  You can edit the text before executing
    the SQL command.

    Once you are familiar with this sample program, you might want to
    change the sample SQL commands to fit your environment.  For
    example, change the document range to some other numbers which are
    appropriate on your server.  The sample SQL commands are located
    on top of the frmCommandIMS.frm and frmCommandMezz.frm files.
 
    For either type of library, you can constrain your search to a 
    folder.  Click on 'Select folders', then choose a top-level folder.  Check the 'Search
    Subfolder' check box if you want to search documents in a folder and
    all subfolders under that folder.  If you check the 'Show Results in IDM 
    ListView' check box, results will be displayed in IDM listview.  Otherwise,
    they will be displayed in a MS listview.  The IDMDS command window has other
    options such as 'Secure Search', 'Access Domain', 'Access Level', which can be
    used to further constrain the search.

    You can also use the 'ADO Properties' menu to set general properties 
    for the ADO command.  For example, you can apply a search limit by specifying
    a value for the "Maximum rows" property.

(6) Click on the 'Execute' button to perform the query.  If the 'Show Results in
    IDM ListView' was checked, the results will be displayed in the
    IDM listview.  If the 'Show Results in IDM ListView' was not checked, a 'Query
    Rowset' window will appear.  You can click the 'ADO Properties' menu
    to view properties or click 'Get Data' button to view rows of
    documents.  Close this window before performing another query.
