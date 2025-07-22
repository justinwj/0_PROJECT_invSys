VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLayoutAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsLayoutAlgorithm
' defines what a layout algorithm must do (e.g. expose a Layout(items, cfg) As Collection method).
Option Explicit

'=== Class: clsLayoutAlgorithm (interface)

' Defines the contract for layout algorithms
Public Function Layout(ByVal items As Collection, _
                       ByVal cfg As clsDiagramConfig) As Collection
    Err.Raise vbObjectError + 3000, "clsLayoutAlgorithm", _
              "Layout must be implemented by concrete algorithms."
End Function
