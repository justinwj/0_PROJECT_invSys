Attribute VB_Name = "modDiagramMaps"
' modDiagramMaps
Option Explicit

' Module: modDiagramMaps
' Parses VBA modules and procedures into diagram items
' Applies moduleFilter and procFilter, instantiates clsDiagramItem objects,
' populates their properties, and returns a Collection of items.

Public Function ParseAndMap(wb As Workbook, moduleFilter As String, procFilter As String) As Collection
    Dim items As New Collection
    Dim vbComp As Object  ' VBIDE.VBComponent
    Dim codeMod As Object ' VBIDE.CodeModule
    Dim lineIndex As Long
    Dim procName As String
    Dim numLines As Long
    Dim item As clsDiagramItem
    Dim startLine As Long

    ' Ensure VBIDE reference
    On Error Resume Next
    For Each vbComp In wb.VBProject.VBComponents
        ' Apply module filter (supports "all mods" wildcard)
        If LCase(moduleFilter) = "all mods" Or LCase(vbComp.Name) Like LCase(moduleFilter) Then
            Set codeMod = vbComp.CodeModule
            numLines = codeMod.CountOfLines
            ' Iterate through all lines to find procedures
            For lineIndex = 1 To numLines
                If codeMod.ProcStartLine(codeMod.ProcOfLine(lineIndex, 0), 0) = lineIndex Then
                    procName = codeMod.ProcOfLine(lineIndex, 0)
                    ' Apply proc filter (supports "all procs" wildcard)
                    If LCase(procFilter) = "all procs" Or LCase(procName) Like LCase(procFilter) Then
                        ' Instantiate a new diagram item
                        Set item = New clsDiagramItem
                        With item
                            .StencilNameU = vbComp.Name    ' Use module name as stencil key by default
                            .LabelText = procName
                            .PosX = 0                      ' TODO: compute X position
                            .PosY = 0                      ' TODO: compute Y position
                        End With
                        items.Add item
                    End If
                    ' Skip to end of this procedure to avoid duplicates
                    startLine = codeMod.ProcStartLine(procName, 0) + _
                                codeMod.ProcCountLines(procName, 0)
                    lineIndex = startLine
                End If
            Next lineIndex
        End If
    Next vbComp
    On Error GoTo 0

    Set ParseAndMap = items
End Function



