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
Option Explicit

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

