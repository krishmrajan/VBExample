This sample shows how to work with the Local DB.

The following areas are covered by this sample:

* Displaying records and groups in the local database.
* Adding an record into the local database.
* Adding a group into the local database.
* Displaying and modifying the properties of a local database record.
* Displaying and modifying the properties of a local group.
* Removing local database records and groups.

This sample will be a useful tool as you develop your own applications. As you are testing code, records in the local database will accumulate. If you create a run-time version of this application (localdb_tool.exe), you can easily clean up the database after testing.


Instructions:

1. Open the "localdb_tool.vbp" project file and run the program.

2. The "Refresh" button causes the list of files in the local database to be updated. The format of each column is:

	a) A red checkmark if the file is checked out
	b) The title of the document
	c) The full path to the local document
	d) The library where the document originated
	e) The document id in the originating library
	f) The version of the document
	g) The checkout date
	h) The user who checked out the document

3. The Add button allows you to manually add a record into the local database. Select a file to add with the open file dialog, then enter the values for the record. Save when you are done.

4. The Properties button allows you to modify a record in the local database. Some of the properties in the local database record are read-only .. these properties are disabled. Save or Cancel when you are done. Refresh causes the  current values for the record to be reloaded into the form.

You can also double-click on any row to edit its properties.

5. The Delete button removes one or more records from the local database. You can select multiple rows using the normal keyboard modifiers (Shift and Control). When you press Delete, you'll be asked to confirm the deletion.

6. To remove all the entries in the local database, use the Delete All button.

7. Exit does what you'd expect.

Please review the code for more information. The comments will explain more about the local database.

