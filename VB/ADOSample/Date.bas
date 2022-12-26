Attribute VB_Name = "dbMzDate"
Option Explicit

Public Function MezDate(inDate As Date) As String
    MezDate = Format(inDate, "yyyy-mm-dd") & "T" & _
    Format(inDate, "hh:mm:ss") & ",000"
End Function


