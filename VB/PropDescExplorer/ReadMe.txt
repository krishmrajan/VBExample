This sample shows how to work with Property Descriptions:
    - It shows how to retrieve property descriptions from a 
      library using the FilterPropertyDescriptions method.
    - It shows how to call GetState to find out about a PropertyDescription.
    - It shows how to call FormatValue to get a text representation
      for a property.
	- It shows how to retrieve classDescriptions using the 
	  FilterClassDescription method.
	  
Since it shows all of the state information about property descriptions,
it is also useful to explore the property descriptions available on a library.

Instructions: 
1.  From Visual Basic, load the "PropDescExplorer.vbp" project file,
    and run the program. 
2.  Expand a library to browse.  You may be prompted to log on.
3.  Expand a class type, for example "Document Classes."  You will see
    under the class type a list of classes, and of all property descriptions
	for that class type.
4.  You may expand a class to see what property descriptions are available
    for the class.
5.  To show the details of a property description, select it in the treeview.
6.  If you would like to see the choices available, select the 
    "Show Choices (first 1000 only)" checkbox.  This will show the choices
    available for the property description.  To save time, only the first
    1000 choices are shown.
	While you can change the other controls (states, default value, etc.),
    this has no effect on the property description.  The controls are enabled
    instead of disabled to make them easier to read.
