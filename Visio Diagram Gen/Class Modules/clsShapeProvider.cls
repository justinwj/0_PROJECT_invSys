VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShapeProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsShapeProvider
' Purpose: Encapsulates loading and caching of Visio stencil masters.
' Requires references:
'   - Microsoft Visio XX.0 Type Library
'   - Microsoft Scripting Runtime

Option Explicit

Private VisioApp As Visio.application
Private stencilCache As Scripting.Dictionary

' Initialize must be called once, passing the Visio Application object
Public Sub Initialize(ByVal application As Visio.application)
    Set VisioApp = application
    Set stencilCache = New Scripting.Dictionary
End Sub

' Returns the ID of the specified master in the given stencil
Public Function GetMasterID(ByVal stencilPath As String, _
                            ByVal masterName As String) As Long
    Dim m As Visio.master
    Set m = GetMaster(stencilPath, masterName)
    GetMasterID = m.ID
End Function

' Returns the Master object by name, loading the stencil if needed
Public Function GetMaster(ByVal stencilPath As String, _
                          ByVal masterName As String) As Visio.master
    Dim stencilDoc As Visio.Document
    Set stencilDoc = LoadStencil(stencilPath)
    Dim m As Visio.master
    For Each m In stencilDoc.Masters
        If StrComp(m.NameU, masterName, vbTextCompare) = 0 Then
            Set GetMaster = m
            Exit Function
        End If
    Next m
    Err.Raise vbObjectError + 1000, "clsShapeProvider", _
              "Master '" & masterName & "' not found in stencil '" & stencilPath & "'"
End Function

' Loads the stencil document and caches it; returns the Visio.Document
Private Function LoadStencil(ByVal stencilPath As String) As Visio.Document
    If Not stencilCache.Exists(stencilPath) Then
        Dim doc As Visio.Document
        ' Open stencil docked; adjust flags as needed
        Set doc = VisioApp.Documents.OpenEx(stencilPath, visOpenDocked)
        stencilCache.Add stencilPath, doc
    End If
    Set LoadStencil = stencilCache(stencilPath)
End Function

