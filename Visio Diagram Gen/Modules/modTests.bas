Attribute VB_Name = "modTests"
' modTests
Option Explicit

' clsCallSites, clsCallSiteMapProvider

' Test for clsDiagramBuilder.BuildConnections
Public Sub TestDiagramBuilder_BuildConnections()
    Dim builder      As clsDiagramBuilder
    Dim fakeApp      As Object
    Dim fakeDoc      As Object
    Dim fakePage     As Object
    Dim fakeShapes   As Object
    Dim connection   As Variant
    Dim connections  As Collection
    Dim connShapes   As Collection
    Dim connector    As Object

    ' Arrange fake Visio environment
    Set fakeApp = CreateObject("Scripting.Dictionary")
    Set fakeDoc = CreateObject("Scripting.Dictionary")
    Set fakePage = CreateObject("Scripting.Dictionary")
    Set fakeShapes = CreateObject("Scripting.Dictionary")

    ' Fake page.Shapes.ItemFromID returns a shape dict with ID key
    fakeShapes.Add 1001, CreateObject("Scripting.Dictionary"): fakeShapes(1001).Add "ID", 1001
    fakeShapes.Add 2002, CreateObject("Scripting.Dictionary"): fakeShapes(2002).Add "ID", 2002
    fakePage.Add "Shapes", fakeShapes
    fakeDoc.Add "Pages", CreateObject("Scripting.Dictionary"): fakeDoc("Pages").Add 1, fakePage

    ' Arrange connections
    Set connections = New Collection
    connections.Add Array(1001, 2002)

    ' Act
    Set builder = New clsDiagramBuilder
    Set builder.Application = fakeApp
    Set builder.Document = fakeDoc
    Set connShapes = builder.BuildConnections(connections)

    ' Assert: one connector created
    If connShapes.Count <> 1 Then Err.Raise vbObjectError + 520, _
        "TestDiagramBuilder_BuildConnections", "Expected 1 connector, got " & connShapes.Count

    Debug.Print "TestDiagramBuilder_BuildConnections passed"
End Sub

' Test for scanning modules
Public Sub TestLoadAllCallSites()
    Dim calls As Collection
    Set calls = LoadAllCallSites()
    Debug.Print "Total call-sites found: " & calls.Count
    ' Optionally add assertions based on known codebase
End Sub

' Test for clsCallSite.GetID
Public Sub TestCallSite_GetID()
    Dim cs As clsCallSite
    Set cs = New clsCallSite
    cs.CallerModule = "ModuleA"
    cs.CallerProc = "Proc1"
    cs.CalleeModule = "ModuleB"
    cs.CalleeProc = "Proc2"

    If cs.GetID <> "ModuleA.Proc1->ModuleB.Proc2" Then
        Err.Raise vbObjectError + 513, "TestCallSite_GetID", _
            "Expected ID 'ModuleA.Proc1->ModuleB.Proc2', got '" & cs.GetID & "'"
    End If

    Debug.Print "TestCallSite_GetID passed"
End Sub

' Test for clsCallSiteMapProvider.MapCallSites
Public Sub TestCallSiteMapProvider()
    Dim sites As Collection
    Set sites = New Collection
    Dim cs As clsCallSite
    Dim fakeCaller As Object
    Dim fakeCallee As Object
    Dim shapesDict As Scripting.Dictionary
    Dim provider As clsCallSiteMapProvider
    Dim connections As Collection
    Dim conn As Variant

    ' Arrange: create a call-site
    Set cs = New clsCallSite
    cs.CallerModule = "ModuleA"
    cs.CallerProc = "Proc1"
    cs.CalleeModule = "ModuleB"
    cs.CalleeProc = "Proc2"
    sites.Add cs

    ' Arrange: fake shapes with IDs using dictionary
    Set fakeCaller = CreateObject("Scripting.Dictionary")
    fakeCaller.Add "ID", 1001
    Set fakeCallee = CreateObject("Scripting.Dictionary")
    fakeCallee.Add "ID", 2002

    Set shapesDict = New Scripting.Dictionary
    shapesDict.Add "ModuleA.Proc1", fakeCaller
    shapesDict.Add "ModuleB.Proc2", fakeCallee

    ' Act
    Set provider = New clsCallSiteMapProvider
    Set connections = provider.MapCallSites(sites, shapesDict)

    ' Assert: one connection returned
    If connections.Count <> 1 Then
        Err.Raise vbObjectError + 514, "TestCallSiteMapProvider", _
            "Expected 1 connection, got " & connections.Count
    End If

    ' Assert: correct IDs
    conn = connections(1)
    If conn(0) <> fakeCaller("ID") Or conn(1) <> fakeCallee("ID") Then
        Err.Raise vbObjectError + 515, "TestCallSiteMapProvider", _
            "Expected connection (" & fakeCaller("ID") & "," & fakeCallee("ID") & "), got (" & conn(0) & "," & conn(1) & ")"
    End If

    Debug.Print "TestCallSiteMapProvider passed"
End Sub

' PASSED TESTS clsDiagramConnection
Public Sub TestDrawConnections()
    Dim items As New Collection
    Dim conns As New Collection
    Dim it As clsDiagramItem
    Dim connObj As clsDiagramConnection
    Dim visApp As Object
    Dim visDoc As Object
    Dim visPage As Object
    Dim rectMaster As Object
    Dim stencil As Object
    Dim shp As Object

    ' Initialize Visio
    On Error Resume Next
    Set visApp = CreateObject("Visio.Application")
    On Error GoTo 0
    If visApp Is Nothing Then
        MsgBox "Visio not available.", vbCritical
        Exit Sub
    End If
    visApp.Visible = True

    ' Create a new Visio document
    Set visDoc = visApp.Documents.Add("")
    If visDoc Is Nothing Then
        MsgBox "Unable to create a Visio document.", vbCritical
        Exit Sub
    End If

    ' Ensure at least one page exists
    If visDoc.Pages.Count = 0 Then visDoc.Pages.Add
    Set visPage = visDoc.Pages(1)

    ' Open Basic Shapes stencil to get Rectangle master
    On Error Resume Next
    Set stencil = visApp.Documents.OpenEx("Basic Shapes.vssx", 64)
    If stencil Is Nothing Then Set stencil = visApp.Documents.OpenEx("Basic Shapes.vss", 64)
    On Error GoTo 0
    If stencil Is Nothing Then
        MsgBox "Unable to open Basic Shapes stencil.", vbCritical
        Exit Sub
    End If
    Set rectMaster = stencil.Masters("Rectangle")
    If rectMaster Is Nothing Then
        MsgBox "Master 'Rectangle' not found in stencil.", vbCritical
        Exit Sub
    End If

    ' Prepare test items
    Set it = New clsDiagramItem: it.LabelText = "A": it.PosX = 1: it.PosY = 5: items.Add it
    Set it = New clsDiagramItem: it.LabelText = "B": it.PosX = 3: it.PosY = 5: items.Add it

    ' Drop shapes on new drawing
    For Each it In items
        Set shp = visPage.Drop(rectMaster, it.PosX, it.PosY)
        shp.Text = it.LabelText
        shp.NameU = it.LabelText
    Next it

    ' Draw connection between shapes
    Set connObj = New clsDiagramConnection
    connObj.FromID = "A": connObj.ToID = "B": conns.Add connObj
    DrawConnections items, conns

    MsgBox "Diagram generated successfully.", vbInformation
End Sub

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

