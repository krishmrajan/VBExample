This is a low-level sample program which demonstrates how to use a recognizer plugin.  This sample is
provided to show programming techniques using CDRecognizer objects and IDM Objects.


The purpose of a this project is to create and test a 3rd party recognizer within the Panagon desktop environment.  

Recognizer Setup:

Steps to set up plug-in:

1. Create the dll: Open the VB project and select File...Make TextRecognizer.dll.  This will create a recognizer dll within the project directory.

2. Register the dll

3. Manually create the Registry Key:
	
   Open the following path within the registry using regedit	

	HKEY_LOCAL_MACHINE\SOFTWARE\FileNET\IDM\Recognizers

There can be found three keys already existing numbered 1-3.  These are the OLE Recognizers that come with Panagon.  Create a fourth Key named "4". Within the new Key, change the value of the substring called default from [value not set] to "QA.TextRecognizer" which is the ProgID of the registered dll. Optionally, create another substring called "Name" and modify its value with a string in order to give the recognizer a name.

4. VB projects using the text recognizer either need to have a reference to the registered dll or the sample VB project must be added as part of
a group project.

Recognizer Usage:
   
   NOTE: See sample parent and child text files within this directory. 

This recognizer plug-in is designed to work specifically with Text documents and finds formatted text within a "parent" document in order to find paths to "child" documents.  Once the recognizer's filepath is set to that of a text file it can then be used to search for links within that text document.  The format for a valid link to a child consists of a line of content within the parent which conforms to the following format:

      C:\Child1.txt
      C:\Temp\Child2.txt

In the above example, if these two lines appeared within the parent then two links would have been found.  A path is the mimimum requirement to qualify a line of text as a link.  The recognizer will set default properties for each link.

   If the user wishes to set link properties within the parent document then the following format must be used for each line:

	Path | ClassID(values 0 or 1) | Link Uses Relative Path (values 0 - 5) | Update Mode (values 0 to 2)

	example:
	
	C:\Child1.txt|1|2|0

	..would mean a child located at the given filepath which is a Dynamic link, which uses the Absolute File path, and has an update mode of Automatic

The value for each property, except the filepath, corresponds to the literal enumeration value for each property within the Panagon API. Any line of text which does not have a colon (:) as its second character will be treated as regular content by the recognizer and will not be processed as a link.

WARNING: if the above pattern is not followed or invalid values are passed as property values, the recognizer may fail to work properly.






