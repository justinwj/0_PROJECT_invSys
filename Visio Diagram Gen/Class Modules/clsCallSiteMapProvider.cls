VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCallSiteMapProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCallSiteMapProvider
' Purpose: Concrete clsMapProvider for testing: returns one DiagramItem
Option Explicit

Implements clsMapProvider

' Forwarding method for late-bound calls
Public Function GetItems(ByVal wb As Workbook, ByVal cfg As clsDiagramConfig) As Collection
    Set GetItems = clsMapProvider_GetItems(wb, cfg)
End Function

' Implementation of clsMapProvider.GetItems
Public Function clsMapProvider_GetItems(ByVal wb As Workbook, _
                                      ByVal cfg As clsDiagramConfig) As Collection
    Dim items As New Collection
    Dim item As clsDiagramItem
    Dim meta As clsMasterMeta
    
    ' Load stub metadata for Rectangle master
    Set meta = LoadStencilMasterMetadata(modDiagramCore.ShapeProvider, "Basic_U.vssx", "Rectangle")
    
    ' Create a diagram item at the origin
    Set item = New clsDiagramItem
    Set item.meta = meta
    item.PosX = cfg.OriginX
    item.PosY = cfg.OriginY
    
    items.Add item
    Set clsMapProvider_GetItems = items
End Function

' Other methods in clsCallSiteMapProvider remain as-is (e.g., MapCallSites)

' Maps a collection of clsCallSite to connector pairs (fromID,toID)
Public Function MapCallSites(sites As Collection, shapesDict As Scripting.Dictionary) As Collection
    Dim connections As New Collection
    Dim site As clsCallSite
    Dim shapeCaller As Object
    Dim shapeCallee As Object
    Dim idCaller As Long
    Dim idCallee As Long

    For Each site In sites
        Set shapeCaller = shapesDict(site.CallerModule & "." & site.CallerProc)
        Set shapeCallee = shapesDict(site.CalleeModule & "." & site.CalleeProc)

        On Error Resume Next
        idCaller = shapeCaller.ID
        If Err.Number <> 0 Then
            Err.Clear
            idCaller = shapeCaller("ID")
        End If

        Err.Clear
        idCallee = shapeCallee.ID
        If Err.Number <> 0 Then
            Err.Clear
            idCallee = shapeCallee("ID")
        End If
        On Error GoTo 0

        connections.Add Array(idCaller, idCallee)
    Next site

    Set MapCallSites = connections
End Function

