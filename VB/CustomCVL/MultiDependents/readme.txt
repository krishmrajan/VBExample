This sample program demonstrates how to create a CVL 
plugin that implements dependent CVLs. In this example,
the value selected from a CVL determines the list of 
choices for the other CVLs. Specifically The country 
selected determines the list of states. The state 
selected determines the list of cities. This data is 
stored in an Microsft Access database named CustomCVL.mdb

This sample works on desktop and web server applications, 
although there are differences in how the code works for 
each type of application. For example, on thin clients, 
the dependency behavior is not supported. 

Instructions:

To install the plug-in, you must be an administrator on the
DS Library.  Use the DocType and CVL Admin Tool to add the 
CVL Plugin to the library:

   1. Create a CVL Plugin named Whatever. For its data, add 
      the Class from the dll: MultiDependents.Geography. 
      MultiDependents.Geography is the name of this project 
      and the class that implements 
      IDMObjects.IFnPropertyDescriptionPlugin

   2. Create a Doc Class as follows:

                 Doc Class Name = MyDocClass

                PropertyName | CVL Name
                ------------ | ---------
      (Optional)Title        |
      (Country) SVCP Str     |  <PI>Whatever<\PI>
      (State)   SVCP Str CVL |  <PI>Whatever<\PI>
      (City)    SVCP Str CVL2|  <PI>Whatever<\PI>

   Note: "SVCP Str *" are just some names here. You can modify them 
   as you wish.

   3. On Client:
      1. Register the Dll using regsvr32.exe or use VB to 
         recompile the project
      2. CustomCVL.mdb must exist in the same folder where the 
         dll is registered
 