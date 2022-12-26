Attribute VB_Name = "Globals"
Public goDSLib As New IDMObjects.Library
Public gbDSLogOff As Boolean
Public goISLib As New IDMObjects.Library
Public gbISLogOff As Boolean
Public goPersist As New clsRegPersist
Public goPropDescs As IDMObjects.PropertyDescriptions
Public gfSettings As New frmSettings
Public gcHeadings As Collection
Public gcPropNames As Collection
Public Const gsAppName = "IDM HRDemo"
Public Const gsSectionName = "LibSettings"

