VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMapProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsMapProvider
' Class: clsMapProvider
' Purpose: Base interface for all diagram map providers.
'          Providers scan a source (e.g., VBA modules) and produce a collection of clsDiagramItem.
Option Explicit
' Note: Concrete subclasses must implement GetItems.
Public Function GetItems(ByVal wb As Workbook, ByVal cfg As clsDiagramConfig) As Collection
    ' Return a Collection of clsDiagramItem objects
    Err.Raise vbObjectError + 2000, "clsMapProvider", _
              "GetItems must be implemented by concrete map providers."
End Function

' Optional: Allows a provider to register itself with a central registry
' Public Sub Register()
'     modDiagramMaps.RegisterProvider Me
' End Sub

