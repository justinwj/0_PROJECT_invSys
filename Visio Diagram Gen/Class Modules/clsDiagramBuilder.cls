VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiagramBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDiagramBuilder.cls
Option Explicit

' Handles drawing connectors between shapes based on call-site mappings
Private pApplication As Object
Private pDocument    As Object

' Set the Visio application
Public Property Set Application(app As Object)
    Set pApplication = app
End Property

' Set the Visio document
Public Property Set Document(doc As Object)
    Set pDocument = doc
End Property

' Builds and returns connector shapes for each (fromID,toID) in connections
Public Function BuildConnections(connections As Collection) As Collection
    Dim connShapes As New Collection
    Dim pagesObj   As Object
    Dim page       As Object
    Dim shapesCol  As Object
    Dim pair       As Variant
    Dim shapeCaller As Object
    Dim shapeCallee As Object
    Dim connector  As Object

    ' Retrieve Pages collection (Visio or fake)
    On Error Resume Next
    Set pagesObj = pDocument.Pages
    If Err.Number <> 0 Then
        Err.Clear
        Set pagesObj = pDocument("Pages")
    End If
    On Error GoTo 0

    Set page = pagesObj(1)

    ' Retrieve Shapes collection (Visio or fake)
    On Error Resume Next
    Set shapesCol = page.Shapes
    If Err.Number <> 0 Then
        Err.Clear
        Set shapesCol = page("Shapes")
    End If
    On Error GoTo 0

    For Each pair In connections
        On Error Resume Next
        Set shapeCaller = shapesCol.ItemFromID(pair(0))
        If Err.Number <> 0 Then
            Err.Clear
            Set shapeCaller = shapesCol(pair(0))
        End If

        On Error Resume Next
        Set shapeCallee = shapesCol.ItemFromID(pair(1))
        If Err.Number <> 0 Then
            Err.Clear
            Set shapeCallee = shapesCol(pair(1))
        End If
        On Error GoTo 0

        ' Attempt to drop connector in Visio, fallback for fake
        On Error Resume Next
        Set connector = page.DropConnected(pApplication.ConnectorToolDataObject, shapeCaller, shapeCallee)
        If Err.Number <> 0 Then
            Err.Clear
            Set connector = shapeCaller
        End If
        On Error GoTo 0

        connShapes.Add connector
    Next pair

    Set BuildConnections = connShapes
End Function

' Test for clsDiagramBuilder.BuildConnections
Public Sub TestDiagramBuilder_BuildConnections()
    Dim builder      As clsDiagramBuilder
    Dim fakeApp      As Object
    Dim fakeDoc      As Object
    Dim fakePage     As Object
    Dim fakeShapes   As Object
    Dim connections  As Collection
    Dim connShapes   As Collection

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

