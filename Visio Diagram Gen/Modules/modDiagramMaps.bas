Attribute VB_Name = "modDiagramMaps"
' modDiagramMaps
Option Explicit

' Module: modDiagramMaps
' Parses VBA modules and procedures into diagram items
' Applies moduleFilter and procFilter, instantiates clsDiagramItem objects,
' populates their properties, and returns a Collection of items.
' Parses VBA project, applies module and procedure filters,
' and returns a collection of clsDiagramItem for Visio rendering.

' Simplified ParseAndMap stub for testing without VBIDE ProcStartLine dependencies
Public Function ParseAndMap(wb As Workbook, moduleFilter As String, procFilter As String) As Collection
    Dim items As New Collection
    Dim vbComp As Object  ' Late-bound VBComponent
    Dim item As clsDiagramItem

    ' Loop through components matching moduleFilter
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Name Like moduleFilter Then
            ' Create a diagram item per module for stub
            Set item = New clsDiagramItem
            item.StencilNameU = "Rectangle"    ' default shape
            item.LabelText = vbComp.Name
            item.PosX = items.Count * 1#          ' simple horizontal layout
            item.PosY = 0#
            items.Add item
        End If
    Next vbComp

    ' Return collection of diagram items
    Set ParseAndMap = items
End Function

'===== Notes on update =====
' • Removed ProcStartLine/ProcCountLines usage to avoid VBIDE dependencies.
' • Now creates one item per VBComponent matching moduleFilter.
' • This stub ensures ParseAndMap compiles and runs regardless of VBIDE reference.
' • Replace with full implementation when ready to handle procedures and code scanning.


