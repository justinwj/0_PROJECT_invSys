Attribute VB_Name = "modDiagramCore"
' modDiagramCore
Option Explicit

' Ensure you have two class modules: clsMasterMeta and clsDiagramConfig
' clsMasterMeta with public properties: FileName, DisplayNameU, DisplayName, ID, Width, Height, Path, LangCode
' clsDiagramConfig with public properties: DiagramType, ModuleFilter, ProcFilter, ScaleMode, ExportFormat

' === Module-level declarations ===
Public gMasterDict As Object      ' Scripting.Dictionary of clsMasterMeta objects keyed by DisplayNameU
Private gConfig As clsDiagramConfig

' === Master metadata infrastructure ===
'-------------------------------------------------------------------------------
' Load real metadata from the "StencilMasters" worksheet
' Builds gMasterDict of clsMasterMeta objects
'-------------------------------------------------------------------------------
Public Sub LoadStencilMasterMetadata()
    On Error Resume Next
    Call AddRequiredReferences
    On Error GoTo 0

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim key As String
    Dim meta As clsMasterMeta

    Set ws = ThisWorkbook.Worksheets("StencilMasters")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        key = Trim(CStr(ws.Cells(i, 2).Value))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                Set meta = New clsMasterMeta
                With meta
                    .FileName = CStr(ws.Cells(i, 1).Value)
                    .DisplayNameU = key
                    .DisplayName = CStr(ws.Cells(i, 3).Value)
                    .ID = CLng(ws.Cells(i, 4).Value)
                    .Width = CDbl(ws.Cells(i, 5).Value)
                    .Height = CDbl(ws.Cells(i, 6).Value)
                    .Path = CStr(ws.Cells(i, 7).Value)
                    .LangCode = CStr(ws.Cells(i, 8).Value)
                End With
                dict.Add key, meta
            Else
                Debug.Print "Skipping duplicate key: " & key
            End If
        End If
    Next i

    Set gMasterDict = dict
    Debug.Print "LoadStencilMasterMetadata: Loaded " & dict.Count & " unique master(s)."
End Sub

'-------------------------------------------------------------------------------
' Standard module stub renamed to avoid conflict
' Place this stub in modDiagramCore for testing purposes
'-------------------------------------------------------------------------------
Public Function LoadStencilMasterMetadataStub() As Object
    Dim dictMasters As Object
    Set dictMasters = CreateObject("Scripting.Dictionary")
    
    ' TODO: Replace with dynamic loading logic
    Dim meta As clsMasterMeta
    Set meta = New clsMasterMeta
    meta.FileName = "Basic_UML.vssx"
    meta.DisplayNameU = "Basic_UML"
    meta.DisplayName = "Basic UML Shapes"
    meta.ID = 1
    meta.Width = 0
    meta.Height = 0
    meta.Path = "C:\Stencils\Basic_UML.vssx"
    meta.LangCode = "en"
    dictMasters.Add meta.DisplayNameU, meta
    
    Set LoadStencilMasterMetadataStub = dictMasters
End Function

Public Function GetMasterMetadata(ByVal masterNameU As String) As clsMasterMeta
    If gMasterDict Is Nothing Then LoadStencilMasterMetadata
    If gMasterDict.Exists(masterNameU) Then
        Set GetMasterMetadata = gMasterDict(masterNameU)
    Else
        Err.Raise vbObjectError + 513, "GetMasterMetadata", _
            "Master '" & masterNameU & "' not found in metadata."
    End If
End Function

' === Configuration loader ===
' Reads values from the DiagramConfig table into the cfg object
Public Sub LoadDiagramConfig(ByVal cfg As clsDiagramConfig)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rw As ListRow
    Dim key As String, val As Variant
    Set ws = ThisWorkbook.Worksheets("DiagramConfig")
    On Error GoTo ErrHandler
    Set tbl = ws.ListObjects("DiagramConfig")
    For Each rw In tbl.ListRows
        key = UCase(Trim(rw.Range.Cells(1, 1).Value))
        val = rw.Range.Cells(1, 2).Value
        Select Case key
            Case "DIAGRAMTYPE":    cfg.DiagramType = val
            Case "MODULEFILTER":   If Len(val) > 0 Then cfg.moduleFilter = val
            Case "PROCFILTER":     If Len(val) > 0 Then cfg.procFilter = val
            Case "SCALEMODE":      cfg.ScaleMode = val
            Case "EXPORTFORMAT":   cfg.ExportFormat = val
        End Select
    Next rw
    Exit Sub
ErrHandler:
    Debug.Print "Error loading DiagramConfig: ", Err.Description
End Sub

' Factory function to create and return a populated configuration
Public Function GetConfig() As clsDiagramConfig
    Dim cfg As clsDiagramConfig
    Set cfg = New clsDiagramConfig   ' Sets defaults in Class_Initialize
    LoadDiagramConfig cfg            ' Overwrite with table values
    Set GetConfig = cfg              ' Return the instance
End Function

' === Visio environment setup ===
' Placeholder for Visio initialization; avoids compile errors if not yet implemented
Public Sub PrepareVisioEnvironment()
    ' TODO: implement Visio application and document setup
End Sub

' === Main orchestrator ===
' RunDiagramGeneration: full pipeline using config-driven parameters
Public Sub RunDiagramGeneration()
    Dim cfg As clsDiagramConfig
    Dim items As Collection
    Dim result As Variant

    ' 1) Load user-defined settings
    Set cfg = GetConfig()
    Debug.Print "[Diagram] Type=" & cfg.DiagramType & _
                "; ModuleFilter=" & cfg.moduleFilter & _
                "; ProcFilter=" & cfg.procFilter

    ' 2) Parse and map VBA code to Visio stencil directives
    On Error Resume Next
    result = Application.Run("modDiagramMaps.ParseAndMap", _
                             ThisWorkbook, cfg.moduleFilter, cfg.procFilter)
    On Error GoTo 0
    If TypeName(result) = "Collection" Then
        Set items = result
    Else
        Set items = New Collection
        Debug.Print "[Diagram] Warning: no mapped items returned."
    End If

    ' 3) Prepare Visio environment and load stencil masters
    PrepareVisioEnvironment
    LoadStencilMasterMetadata

    ' 4) Render mapped items onto the Visio page
    DrawMappedElements items, cfg.ScaleMode, cfg.ExportFormat

    ' 5) Post-render: apply additional layout (tiling, fitting, etc.)
    ApplyLayout cfg.ScaleMode

    ' 6) Export diagram using configured format
    modDiagramExport.SaveDiagram cfg.ExportFormat

    Debug.Print "[Diagram] Generation complete."
End Sub
' Note: Adjust DrawMappedElements signature to accept config args
' Public Sub DrawMappedElements(ByVal items As Collection, ByVal ScaleMode As String, ByVal ExportFormat As String)
'     ' …render shapes, apply ScaleMode settings, ready for export…
' End Sub

'-------------------------------------------------------------------------------
' DrawMappedElements
' Iterates gMasterDict and drops each master on the active Visio page
'-------------------------------------------------------------------------------
' DrawMappedElements now drops shapes solely from the opened stencil doc
' — ensures visApp, visDoc, and visPage are set
' — opens stencil if not already loaded
' — handles missing master gracefully
Public Sub DrawMappedElements(ByVal items As Collection, ByVal ScaleMode As String, ByVal ExportFormat As String)
    Dim visApp As Object
    Dim visDoc As Object
    Dim visPage As Object
    Dim stencilDoc As Object
    Dim masterShape As Object
    Const stencilName As String = "Basic_U.vssx"
    Dim item As clsDiagramItem

    ' 1) Get or create Visio application
    On Error Resume Next
    Set visApp = GetObject(, "Visio.Application")
    If visApp Is Nothing Then Set visApp = CreateObject("Visio.Application")
    On Error GoTo 0

    ' 2) Ensure a document and page exist
    If visApp.Documents.Count = 0 Then visApp.Documents.Add ""
    Set visDoc = visApp.ActiveDocument
    If visDoc Is Nothing Then Exit Sub  ' safety check
    If visDoc.Pages.Count = 0 Then visDoc.Pages.Add
    Set visPage = visApp.ActivePage

    ' 3) Open stencil for masters if needed
    On Error Resume Next
    Set stencilDoc = visApp.Documents(stencilName)
    If stencilDoc Is Nothing Then
        Set stencilDoc = visApp.Documents.OpenEx(stencilName, 4)
    End If
    On Error GoTo 0

    ' 4) Drop each item
    For Each item In items
        Set masterShape = Nothing
        On Error Resume Next
        Set masterShape = stencilDoc.Masters(item.StencilNameU)
        On Error GoTo 0
        If masterShape Is Nothing Then
            Debug.Print "[Diagram] Warning: master '" & item.StencilNameU & "' not found in stencil."
        Else
            visPage.Drop masterShape, item.PosX, item.PosY
            visPage.Shapes(visPage.Shapes.Count).Text = item.LabelText
        End If
    Next item

    Debug.Print "[Diagram] Dropped " & items.Count & " shapes"
End Sub

' === Layout and scaling ===
Private Sub ApplyLayout(ByVal ScaleMode As String)
    Select Case LCase(ScaleMode)
        Case "fittopage"
            ActivePage.PageSheet.CellsU("Print.PageScale").FormulaU = "1"
            ActiveWindow.PageFit = 2  ' visFitPage
        Case "autotile"
            ' TODO: implement autotile layout
        Case Else
            ' No layout
    End Select
End Sub

' clsDiagramConnection stub implementation for modDiagramCore
' Stub: draws connections between shapes
' DrawConnections: connects shapes in Visio based on item IDs
Public Sub DrawConnections(items As Collection, conns As Collection)
    Dim visApp    As Object
    Dim visPage   As Object
    Dim dictShapes As Object
    Dim item      As clsDiagramItem
    Dim conn      As clsDiagramConnection
    Dim shp       As Object
    Dim shapeFrom As Object
    Dim shapeTo   As Object

    ' Attach to Visio
    On Error Resume Next
    Set visApp = GetObject(, "Visio.Application")
    If visApp Is Nothing Then Set visApp = CreateObject("Visio.Application")
    On Error GoTo 0
    If visApp Is Nothing Then Exit Sub

    ' Use active page
    Set visPage = visApp.ActivePage
    If visPage Is Nothing Then Exit Sub

    ' Map item IDs to shapes
    Set dictShapes = CreateObject("Scripting.Dictionary")
    For Each item In items
        On Error Resume Next
        Set shp = visPage.Shapes(item.LabelText)
        On Error GoTo 0
        If Not shp Is Nothing Then dictShapes.Add item.LabelText, shp
        Set shp = Nothing
    Next item

    ' AutoConnect shapes
    For Each conn In conns
        If dictShapes.Exists(conn.FromID) And dictShapes.Exists(conn.ToID) Then
            Set shapeFrom = dictShapes(conn.FromID)
            Set shapeTo = dictShapes(conn.ToID)
            If Not shapeFrom Is Nothing And Not shapeTo Is Nothing Then
                shapeFrom.AutoConnect shapeTo, 1   ' visAutoConnectDirNone
            End If
        Else
            Debug.Print "DrawConnections: missing shapes for " & conn.FromID & "?" & conn.ToID
        End If
    Next conn
End Sub
