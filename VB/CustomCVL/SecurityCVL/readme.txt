This sample program demonstrates how to create a CVL 
plugin for setting a document's access level. This example
supplies two CVL's to the library, and uses the value 
selected on one to validate the other. 

This sample only works on desktop applications. 

Instructions:

To install the plug-in, you must be an administrator on the
DS Library.  Use the DocType and CVL Admin Tool to add the 
CVL Plugin to the library:

   1. Create a CVL Plugin named SecurityCVL.UserAccess. For 
      its data, add the name of the class from the dll: 
      SecurityCVL.UserAccess

      SecurityCVL.UserAccess is this project's name and the 
      class that implements IDMObjects.PropertyDescriptionPlugIn

   2. Create a regular CVL JobTitles and add the following data:

         Director
         Manager
         Sr. Engineer
         Engineer
         Intern

   3. Create a Doc Class as follows:

      Doc Class Name = SecurityPlugin (Whatever you want)

                PropertyName | CVL Name
                ------------ | ---------
      (Optional)Title        |
                SVCP Str CVL |  JobTitles
                SVCP Str     |  <PI>SecurityPlug<\PI>

      Note: "SVCP Str CVL" and "SVCP Str" are just labels here, 
      you can modify them as needed.
   
