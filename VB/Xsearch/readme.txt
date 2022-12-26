XSEARCH Sample Application

Introduction
This sample application was produced to illustrate one method of performing cross-library searches using the Panagon IDM API.  The sample will perform a query across two libraries.  The libraries may be of similar type (e.g. two IDMIS systems) or dissimilar type (an IDMIS system and an IDMDS system).

Objects / Techniques Utilized
The sample makes heavy use of the following Panagon IDM objects and techniques:
	ADO queries (IDMIS and IDMDS)
	Panagon IDM "Library" objects
	Panagon IDM "PropertyDescription" objects
	Panagon IDM  "ObjectSet" collections

Approach Taken
The approach taken to performing cross-library queries is to run separate ADO queries against each library.  There clearly needs to be a way for identifying analogous fields (properties) in each library.  This appears straightforward in the case of "system" fields in libraries of similar type, for example, the F_ENTRYDATE field in one IDMIS system will be analogous to the F_ENTRYDATE field in another IDMIS system.  However, it becomes more difficult to match F_ENTRYDATE to a field on an IDMDS server -- there are a number of date fields associated with documents on an IDMDS system (e.g. idmDateAdded, idmVerCreateDate).  Also, "user" fields which are known to be analogous, even on systems of similar type, may have been given different names when defined.  For example, the "AccNumber" field on one IDMIS system may correspond to the "AccountNumber" field on another.  Therefore, it becomes virtually impossible to equate library fields automatically.  This sample provides a screen where a user can match fields on each of the two chosen libraries.  These fields can then be used in the application's cross-library search screen.  In production applications, these similarities would be defined by an administrator and saved in a configuration file (or registry entry).  Note that only the field names from the first library are used in the search screen and in the subsequent result listings.

By default, the document identifier (F_DOCNUMBER or idmId) and library name will be displayed for each result.  The application provides another screen where the user can specify the fields required in the results listing.  These display fields may be the ones which have been identified as cross-referenced for the two libraries as well as the ones specific to a particular library.

When a result item is selected within the results list, the application simply launches the document into the appropriate viewer.

It should be noted that a large amount of the code in this application is devoted to handling the matching of field names from the two libraries in search definition and the presentation of the results set.
