Attribute VB_Name = "modTests"
' modTests
Option Explicit


' PASSED TESTS clsDiagramItem
' Stub for testing mapping: returns a test collection of clsDiagramItem instances
Public Function TestParseAndMap(wb As Workbook, moduleFilter As String, procFilter As String) As Collection
    Dim items As New Collection
    Dim it As clsDiagramItem

    ' Example test nodes
    Set it = New clsDiagramItem
    it.StencilNameU = "Ellipse"
    it.LabelText = "Node A"
    it.PosX = 1#
    it.PosY = 1#
    items.Add it

    Set it = New clsDiagramItem
    it.StencilNameU = "Diamond"
    it.LabelText = "Node B"
    it.PosX = 4#
    it.PosY = 2#
    items.Add it

    Set TestParseAndMap = items
End Function

' Wrapper to test the TestParseAndMap stub in modDiagramMaps.bas
Public Sub TestRunParseAndMap()
    Dim items As Collection

    ' 1) Invoke real ParseAndMap implementation
    Set items = ParseAndMap(ThisWorkbook, "*", "*")
    Debug.Print "[Test] Parsed items count: " & items.Count

    ' 2) Prepare Visio and load stencil metadata
    PrepareVisioEnvironment
    LoadStencilMasterMetadata

    ' 3) Render parsed items using existing draw routine
    DrawMappedElements items, "FitToPage", "PNG"
    Debug.Print "[Test] Rendered parsed items"
End Sub

'===== Additional helper: list stencil master names =====
' Run this test to print the first 50 master names from the opened stencil
Public Sub TestListStencilMasters()
    Dim stencilDoc As Object
    Dim visApp As Object
    Dim m As Object
    Dim i As Long
    Const stencilName As String = "Basic_U.vssx"

    ' Attach to Visio and ensure stencil loaded
    On Error Resume Next
    Set visApp = GetObject(, "Visio.Application")
    If visApp Is Nothing Then Set visApp = CreateObject("Visio.Application")
    On Error GoTo 0
    Set stencilDoc = visApp.Documents(stencilName)
    If stencilDoc Is Nothing Then Set stencilDoc = visApp.Documents.OpenEx(stencilName, 4)

    ' List master names
    Debug.Print "Listing first 50 masters in " & stencilName
    i = 0
    For Each m In stencilDoc.Masters
        Debug.Print "  [" & m.Name & "]"
        i = i + 1
        If i >= 50 Then Exit For
    Next m
End Sub
' Stub to test rendering a single clsDiagramItem
Public Sub TestDrawSingleItem()
    Dim item As clsDiagramItem
    Dim items As Collection

    ' Prepare environment and stencils
    PrepareVisioEnvironment
    LoadStencilMasterMetadata

    ' Configure one diagram item (adjust name to match stencil master exactly)
    Set item = New clsDiagramItem
    item.StencilNameU = "Rectangle"   ' use a valid master name from the stencil   ' use exact master name from stencil (no extension)
    item.LabelText = "Test Node"
    item.PosX = 2#
    item.PosY = 3#

    ' Collect and render
    Set items = New Collection
    items.Add item
    DrawMappedElements items, "FitToPage", "PNG"

    Debug.Print "Rendered single test item: " & item.StencilNameU
End Sub

' PASSED TEST
' clsDiagramConfig
' Smoke test to verify config loading only
Public Sub TestLoadConfig()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow

    Set ws = ThisWorkbook.Worksheets("DiagramConfig")
    Set tbl = ws.ListObjects("DiagramConfig")
    
    Dim cfg As clsDiagramConfig
    Set cfg = GetConfig()
    Debug.Print "Type:    ", cfg.DiagramType
    Debug.Print "ModFil:  ", cfg.moduleFilter
    Debug.Print "PrFil:   ", cfg.procFilter
    Debug.Print "Scale:   ", cfg.ScaleMode
    Debug.Print "ExpFmt:  ", cfg.ExportFormat
    Debug.Print "Table found: " & tbl.Name
    For Each rw In tbl.ListRows
        Debug.Print "Row: " & rw.Range.Cells(1, 1).Value & " = " & rw.Range.Cells(1, 2).Value
    Next rw

End Sub

' PASSED TESTS clsMasterMeta
'-------------------------------------------------------------------------------
' Test routine for LoadStencilMasterMetadataStub
' Place this sub in a dedicated test module (e.g., modTest) to keep tests separate
'-------------------------------------------------------------------------------
Public Sub TestLoadStencilMasterMetadataStub()
    Dim dictMasters As Object
    Dim key As Variant
    Dim meta As clsMasterMeta
    
    Set dictMasters = LoadStencilMasterMetadataStub()
    If dictMasters Is Nothing Then
        MsgBox "LoadStencilMasterMetadataStub returned Nothing", vbCritical
        Exit Sub
    End If
    
    Debug.Print "--- Loaded Masters ---"
    For Each key In dictMasters.Keys
        Set meta = dictMasters(key)
        Debug.Print "Key=" & key & ", FileName=" & meta.FileName & _
                    ", DisplayName=" & meta.DisplayName & _
                    ", ID=" & meta.ID & _
                    ", Path=" & meta.Path
    Next key
    Debug.Print "Total masters: " & dictMasters.Count
    MsgBox "Test complete: " & dictMasters.Count & " master(s) loaded.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' Test flow for master metadata + rendering stub
' Place this in modTests or modDiagramCore to verify end-to-end stub integration
'-------------------------------------------------------------------------------
Public Sub TestMasterFlow()
    ' Ensure the master dictionary is loaded
    LoadStencilMasterMetadata
    
    ' Quick check of contents
    If gMasterDict Is Nothing Or gMasterDict.Count = 0 Then
        MsgBox "No masters loaded!", vbCritical, "Master Flow Test"
        Exit Sub
    Else
        MsgBox "Loaded " & gMasterDict.Count & " master(s). Now testing DrawMappedElements.", vbInformation, "Master Flow Test"
    End If
    
    ' Call your existing draw routine (stub or real) to drop shapes
    ' Replace DrawMappedElements with your actual entry point
    Call DrawMappedElements_Sstub
End Sub

'-------------------------------------------------------------------------------
' Minimal stub for DrawMappedElements to confirm invocation
' Modify or replace with your real routine when ready
' Used by TestMasterFlow
'-------------------------------------------------------------------------------
Public Sub DrawMappedElements_Sstub()
    Dim key As Variant
    Dim meta As clsMasterMeta
    
    Debug.Print "--- Drawing Elements Stub ---"
    For Each key In gMasterDict.Keys
        Set meta = gMasterDict(key)
        ' In real code you'd call Visio.Drop meta.ID, meta.PosX, meta.PosY
        Debug.Print "Would drop shape '" & meta.DisplayNameU & "' from file '" & meta.FileName & "'."
    Next key
    MsgBox "DrawMappedElements stub executed for " & gMasterDict.Count & " shape(s).", vbInformation, "Draw Stub"
End Sub

