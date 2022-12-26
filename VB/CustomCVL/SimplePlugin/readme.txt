This sample program demonstrates how to create a simple CVL 
plugin that populates the CVL with values from a database.
This sample works on both web servers and in desktop 
applications. 

Instructions:

To run this sample you must put the Country.mdb database on 
the machine where the demo will run. 

To install the plug-in, you must be an administrator on the
DS Library.
Use the DocType and CVL Admin Tool:

   1. Create a CVL Plugin named Whatever. For its data, add 
      SimplePlugIn.Geography.

      SimplePlugIn.Geography is this project's name and the 
      class that implements IDMObjects.PropertyDescriptionPlugin

   2. Create a Doc Class as follows:

         Doc Class Name = MyDocClass

                PropertyName | CVL Name
                ------------ | ---------
      (Optional)Title        |
             SVCP Str CVL    |  <PI>Whatever<\PI>

      Note: "SVCP Str CVL" is just a label here. It represents a
      Single value property that has a CVL assigned to it.

To run this example on a thin client, this dll only
needs to be registered on the web server. Country.mdb must 
also exist in the same location where this dll is registered.
