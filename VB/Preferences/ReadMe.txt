This sample shows how to work with Preferences.

The following areas are covered by this sample:

* Displaying the value of a preference in your application.
* Displaying and modifying a system preference in your application.
* Defining your own custom preference and using it in your application.
* Exporting and importing preferences in an application.


Instructions:

1. Open the "Preference Sample.vbp" project file and run the program.

2. A message box will appear saying "Hey!" Press OK to continue.

3. The top line shows how you can incorporate a preference value in a label or other UI element in your application.

4. The next line shows how you can modify a system preference in your application, without relying on the IDM Configure application. Pressing the "Reset to Default" button will cause the preference value to be return to its default value.

5. The last line shows how you can add a preference to modify the behavior of your application. In this case, the dialog box that appeared when you startup the application can be turned on and off. Also, the message that is displayed can be modified by entering some different text.

6. Finally, the current user's preferences can be exported and imported using the buttons at the bottom of the window.


Please review the code for more information. The comments will guide you through the process of adding preferences to your application.

