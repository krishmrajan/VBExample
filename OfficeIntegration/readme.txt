This is the source code from the Office 97/2000 integration supplied with IDM DT 3.0.  
It provides the following functionality:
- Supports Applicatin Integration Specific preferences
- Supports Compound Documents
- Supports Insertion \ Updating of Property References

There are two software componenets used for this integration IDMMACROAPI.DLL and IDMMACRO.DLL.  IDMMACRO.DLL depends on the existence of IDMMACROAPI.DLL.  Additionally, there is a idmmacro.rc file located in the \idmmacro\resource directory, that is used to provide a resource file.  A developer must use the resouce complier RC.EXE to create a resource file named IDMMACRO.RES that is used with IDMMACRO.DLL

Instructions to run the project in debug mode: 
1. 	If the idmmacro.dll was installed with the IDM DT install unregister this dll by running REGSRV32.EXE with the /u switch.  The following is an example:
      C:\WINNT\system32\REGSVR32.EXE   /u "C:\Program Files\FileNET\IDM\idmmacro.dll"

2.  From Visual Basic, load the \idmmacroapi\idmmacroapi.vbp project file,and run the program. 

3.  Make the idmmacroapi.dll.  This will also register this file.

4.  From Visual Basic, load the \idmmacro\idmmacro.vbp project file, and run the program.  

5. Open one of the Office Macros (idmwrd8.dot, idmxl8source.xls or, idmpp8source.ppt) and under the open application select Tools | Macro | Visual Basic Editor.  This will activate the VBA IDE.

6. In the VBA IDE select Tools | References and select the reference to the idmmacro.vbp.
7. Close the Open Macro file and save the file.
	For Word - Save as a .dot
	For Excel - Save as a .xla
	For PowerPoint - Save as a .ppa



8. Move the file to the approprite startup directory making sure to remove any Macros previously installed during IDM DT install:
	For Word - <Office Install>\Office\Startup
	For Excel - <Office Install>\Office\XLStart
	For PowerPoint - <FileNET Install>IDM.  In PowerPoint make sure to register this new Add-in (.ppa) through the Tools 	|Add-in 
9. Run the desired Office Application and set breakpoints in the VB IDE to see the functionality in the idmmacro.vbp


Instructions to make the IDMMACRO.DLL:

1. From Visual Basic, load the \idmmacroapi\idmmacroapi.vbp project file,and run the program. 

2.  Make the idmmacroapi.dll.  This will also register this file.

3.  From Visual Basic, load the \idmmacro\idmmacro.vbp project file, and select File | Make idmmacro.dll.  

4. Open one of the Office Macros (idmwrd8.dot, idmxl8source.xls or, idmpp8source.ppt) and under the open application select Tools | Macro | Visual Basic Editor.  This will activate the VBA IDE.

5. In the VBA IDE select Tools | References and select the reference to the idmmacro.vbp.
6. Close the Open Macro file and save the file.
	For Word - Save as a .dot
	For Excel - Save as a .xla
	For PowerPoint - Save as a .ppa


7. Move the file to the approprite startup directory making sure to remove any Macros previously installed during IDM DT install:
	For Word - <Office Install>\Office\Startup
	For Excel - <Office Install>\Office\XLStart
	For PowerPoint - <FileNET Install>IDM.  In PowerPoint make sure to register this new Add-in (.ppa) through the Tools 		|Add-in 



Note: If you see a message "Errors during load. refer to '...frmPropertyMgr.log' for detail", Open the form property, select Icon, and change (Icon) to (None). You should not see the message again.