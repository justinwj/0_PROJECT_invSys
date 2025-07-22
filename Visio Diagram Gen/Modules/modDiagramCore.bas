Attribute VB_Name = "modDiagramCore"
' modDiagramCore
Option Explicit

' Ensure you have two class modules: clsMasterMeta and clsDiagramConfig
' clsMasterMeta with public properties: FileName, DisplayNameU, DisplayName, ID, Width, Height, Path, LangCode
' clsDiagramConfig with public properties: DiagramType, ModuleFilter, ProcFilter, ScaleMode, ExportFormat

' === Module-level declarations ===
Public gMasterDict As Object      ' Scripting.Dictionary of clsMasterMeta objects keyed by DisplayNameU
Private gConfig As clsDiagramConfig
' Module-level globals
Private gVisioApp As Visio.application
Private gWorkbook As Workbook
Private gShapeProvider As clsShapeProvider

' Initialize core services: Visio app, workbook, config, and shape providers
Public Sub Initialize(ByVal visApp As Visio.application, _
                      ByVal wb As Workbook, _
                      ByVal cfg As clsDiagramConfig)
    Set gVisioApp = visApp
    Set gWorkbook = wb
    Set gConfig = cfg

    ' Initialize the shape provider
    Set gShapeProvider = New clsShapeProvider
    gShapeProvider.Initialize gVisioApp

    ' Register map providers
    modDiagramMaps.ClearProviders
    Dim prov As Object
    Set prov = New clsCallSiteMapProvider
    modDiagramMaps.RegisterProvider prov
End Sub

' Accessor for Visio.Application
Public Property Get VisioApp() As Visio.application
    Set VisioApp = gVisioApp
End Property

' Accessor for Workbook context
Public Property Get Workbook() As Workbook
    Set Workbook = gWorkbook
End Property

' Accessor for DiagramConfig
Public Property Get Config() As clsDiagramConfig
    Set Config = gConfig
End Property

' Accessor for ShapeProvider
Public Property Get ShapeProvider() As clsShapeProvider
    Set ShapeProvider = gShapeProvider
End Property

Public Sub Generate()
    Dim shapes     As Collection
    Dim layoutAlg As clsLayoutAlgorithm
    Dim shapedDict As Scripting.Dictionary
    Dim connectors As Collection
    Dim drawingDoc As Visio.Document
    Dim pg         As Visio.page
    Dim builder    As clsDiagramBuilder

    ' 1) Gather shape items using gConfig
    Set shapes = modDiagramMaps.ExecuteProviders(gWorkbook, gConfig)
    '                                   ? gConfig, not cfg        :contentReference[oaicite:0]{index=0}

    ' 2) Layout them
    Set layoutAlg = New clsVerticalLayoutAlgorithm
    Set shapes = layoutAlg.Layout(shapes, gConfig) ' :contentReference[oaicite:1]{index=1}

    ' 3) New Visio page
    Set drawingDoc = gVisioApp.Documents.Add("")
    Set pg = drawingDoc.Pages(1)

    ' 4) Draw shapes
    Set builder = New clsDiagramBuilder
    builder.Initialize gVisioApp, pg
    Set shapedDict = builder.DrawItems(shapes, gShapeProvider)

    ' 5) (Stub) connectors
    Set connectors = New Collection
    If connectors.Count > 0 Then
        builder.DrawConnections connectors, shapedDict
    Else
        Debug.Print "No connectors to draw"
    End If

    Debug.Print "Generate: completed " & shapes.Count & " shapes and " & connectors.Count & " connectors."
End Sub

' === Master metadata infrastructure ===
'-------------------------------------------------------------------------------
' Load real metadata from the "StencilMasters" worksheet
' Builds gMasterDict of clsMasterMeta objects
'-------------------------------------------------------------------------------
' Version one: load metadata from the "StencilMasters" sheet
Public Sub LoadStencilMasterMetadataFromWorksheet()
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
    Debug.Print "LoadStencilMasterMetadataFromWorksheet: Loaded " & dict.Count & " unique master(s)."
End Sub

Public Function GetMasterMetadata(ByVal masterNameU As String) As clsMasterMeta
    If gMasterDict Is Nothing Then LoadStencilMasterMetadataFromWorksheet
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
    result = application.Run("modDiagramMaps.ParseAndMap", _
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
    LoadStencilMasterMetadataFromWorksheet

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
            visPage.shapes(visPage.shapes.Count).Text = item.LabelText
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

'--- Stub for loading metadata in modDiagramCore ---
' Version two: stub for testing via clsShapeProvider, with guaranteed active document/page
Public Function LoadStencilMasterMetadata(ByVal provider As clsShapeProvider, _
                                         ByVal stencilPath As String, _
                                         ByVal masterNameU As String) As clsMasterMeta
    Dim meta As clsMasterMeta
    Dim m As Visio.master
    Dim app As Visio.application
    Dim visDoc As Visio.Document
    Dim pg As Visio.page
    Dim shp As Visio.Shape
    Set meta = New clsMasterMeta

    ' Retrieve Visio.Master and metadata
    Set m = provider.GetMaster(stencilPath, masterNameU)
    Set app = m.Document.application

    ' Ensure at least one document is open
    If app.Documents.Count = 0 Then
        Set visDoc = app.Documents.Add("")
    Else
        Set visDoc = app.ActiveDocument
    End If

    ' Ensure at least one page exists
    If visDoc.Pages.Count = 0 Then visDoc.Pages.Add
    Set pg = visDoc.Pages(1)   ' Use the first page explicitly

    ' Populate metadata properties
    meta.FileName = stencilPath
    meta.DisplayNameU = masterNameU
    meta.DisplayName = masterNameU
    meta.ID = m.ID

    ' Drop the shape and capture size
    On Error Resume Next
    Set shp = pg.Drop(m, 0, 0)
    If Not shp Is Nothing Then
        meta.Width = shp.CellsU("Width").ResultIU
        meta.Height = shp.CellsU("Height").ResultIU
        shp.Delete
    End If
    On Error GoTo 0

    meta.Path = stencilPath

    Set LoadStencilMasterMetadata = meta
End Function

'--- Test harness for full pipeline ---
Public Sub TestGenerate()
    Dim visApp As Visio.application
    On Error Resume Next
    Set visApp = GetObject(, "Visio.Application")
    If visApp Is Nothing Then Set visApp = New Visio.application
    On Error GoTo 0

    Dim cfg As clsDiagramConfig
    Set cfg = New clsDiagramConfig
    ' Optional: customize start point and spacing
    cfg.OriginX = 2
    cfg.OriginY = 8
    cfg.VerticalSpacing = 1.5

    ' Initialize and run
    Initialize visApp, ThisWorkbook, cfg
    Generate

    ' Report results
    Debug.Print "TestGenerate: Completed pipeline"
    Debug.Print "Shapes on page: " & visApp.ActivePage.shapes.Count
End Sub
