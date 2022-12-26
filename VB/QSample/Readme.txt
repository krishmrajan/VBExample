The QSample Visual Basic sample program shows how to manipulate objects relating 
to queues.  It also illustrates use of the LogServer object for logging 
and tracing, as well as use of the LocalCache and ServerCache classes. 

This sample is also useful to create queues and queue workspaces required
for all other queue operations.  Since this is usually best done via an
interactive manual session, this sample application provides a complete
basic application to perform these queue maintenance operations, as well
as others.

Instructions:
1. From Visual Basic, load the QSample project file, and run the program. 
2. Select an IDMIS library, and logon. 
3. You must now select values for 'Workspace' and 'Queue name', unless
   doing creation, modification, or deletion of queues or queue workspaces
   (see #8, below).
4. Once you have specified these values, you can examine the contents 
   of the queue by clicking on the 'Start Query' button.  More of the
   queue can be viewed using the 'More' button, while the entire queue
   can be viewed using the 'All' button.  Use the 'Stop' button to stop
   the current query, freeing up resources on the client and on the server.
   Insertion of a new queue entry can also be done via the 'Insert' button
   once a query has been started.
5. If you select a row in the queue contents grid control, the buttons 
   at the bottom of the form are enabled.  These will enable you to add
   a new queue entry, modify an existing one, or delete one.  You can 
   also change the 'busy' state of a selected queue entry. 
6. If your queue definition (queue schema) includes a document id field, 
   the sample application will perform additional functions.  At the time 
   the grid is populated with the queue contents, the application will use the 
   ServerCache object to request migration of any referenced documents.  When
   a row in the grid is selected, the application will use the LocalCache 
   object to prefetch (to the workstation) the referenced document.
7. To be more selective on which queue entries are retreived, use the Query
   button to allow setting filter conditions, then click on the 'Start Query'
   button again.  Only the non-empty fields are used in the query.  When a
   new queue is selected, the filters are cleared. 
8. Creation, modification, and deletion of queues and queue workspaces are
   done via the menu items on the form.  Deletion and modification are done
   on the currently selected item.  The queue creation operation is only
   enabled after a queue workspace has been selected.

Programming Notes: 
This application uses most of the objects associated with queues: Queue, 
QueueEntry, QueueBrowseSet, QueueQuerySpecification, QueueWorkspace and the
various Property-related objects. 

The qMaintform contains most of the logic for viewing queue contents, 
manipulating the contents grid control, and for toggling the busy state of 
a particular queue entry.  The viewing of queue contents is accomplished through
the use of the BrowseSet object, which is constructed in the BuildBrowseSet sub.
Iterating through the BrowseSet, examining Properties and values, and populating
the grid control are accomplished in the DisplayQContents sub.  The function 
FetchQueueEntry demonstrates how to get a particular entry by entry ID. The 
sample application keeps a collection of QueueEntries (cQueueEntries) - this is
emptied in the InitializeContent routine and is populated in DisplayQContents. 
Care must be taken to refresh this collection whenever the queue contents are 
changed. 

The EditEntryForm handles modification of property values for the selected queue
entry.  The edit form, itself, is built dynamically using VB's support for 
control arrays.  The logic in this area primarily deals with form and control 
behavior, and iteration through property collections; interaction with any of 
the queue related objects is minimal. 

The QueryForm handles changes to the filters in the QueueQuerySpecification.
Much of the query form, itself, is built dynamically using VB's support for
control arrays.  Again, the logic in this area primarily deals with form and
control behavior, and iteration through property collections; interaction with
any of the queue related objects is minimal.

The WSCreate form handles creation and updating queue workspace definitions.
It mainly deals with forms, especially updating the qMaintForm.

The QCreate form handles creation and updating queue definitions.  The main
complexity in this form deals with dynamically creating, maintaining, and
scrolling the rows of property descriptions on the window via control arrays.
It also deals with updating the qMaintForm.

The DelConf form handles confirmation of queue deletions.  It is very simple,
and only exists because of the need for a checkbox (otherwise an InputBox
could be used).

Use of the LocalCache and ServerCache objects is straightforward.  The
InitializeGrid subroutine checks the property descriptions and data types of the
queue to see if there is a document identifier field.  If one is found, its
column position is saved in the iDocIndex variable.  The DisplayQContents
subroutine uses this variable to extract document id values from queue entries
as they are added to the grid control.  These document id values are passed to
the ServerCache object for migration into the server-side cache.  At the point
that an individual row of the grid control is selected (a specific queue entry
is chosen), the grdQueueData_SelChange subroutine will use the LocalCache object
to request migration of the referenced document id.  Again, this is done by
using the iDocIndex global variable, which identifies the grid position
containing the document id.  This is the place in the code that also controls
the enabling of the 'View' button.

The logging and tracing functions are encapsulated in two classes: CErrorLog
and CTraceLog.  At various points in the sample application, trace or error
entries are generated using the FileNet logging facilities.  See the online help
files for information on how to configure your workstation to capture and view
these events.

The following is a list of production-style queue functionality used in
this application:
    opening queues
    using properties and property descriptions
    inserting entries
    deleting entries
    updating entries
    browsing entries
    selecting entries by entry ID
    making an entry read/write
    using filters
    listing workspaces
    listing queues
    using non-filter selectivity of query specifications in browsing,
        fetching, or counting
    clearing a query specification
    sorting (although completely manually specified)

The following is a list of additional maintenance-style queue functionality
used in this application:
    overriding the busy indication
    resetting the busy indication
    browsing all entries, including those that are not normally visible
    counting entries
    creating workspaces
    updating workspace definitions
    deleting workspaces
    creating queues
    updating queue definitions
    deleting queues
    unlimited browsing (not putting a limit on the number of entries retrieved)

The following is a list of production-style queue functionality not used in
this application:
    fetching entries
    use of the QueueEntries object
    canceling a browse
    clearing a queue entry

The following is a list of additional maintenance-style queue functionality
not used in this application:
    emptying a queue
