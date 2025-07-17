VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCallSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCallSite
' Class data for modDiagramCore
Option Explicit

Private pCallerModule As String
Private pCallerProc   As String
Private pCalleeModule As String
Private pCalleeProc   As String

' CallerModule property
Public Property Get CallerModule() As String
    CallerModule = pCallerModule
End Property
Public Property Let CallerModule(ByVal val As String)
    pCallerModule = val
End Property

' CallerProc property
Public Property Get CallerProc() As String
    CallerProc = pCallerProc
End Property
Public Property Let CallerProc(ByVal val As String)
    pCallerProc = val
End Property

' CalleeModule property
Public Property Get CalleeModule() As String
    CalleeModule = pCalleeModule
End Property
Public Property Let CalleeModule(ByVal val As String)
    pCalleeModule = val
End Property

' CalleeProc property
Public Property Get CalleeProc() As String
    CalleeProc = pCalleeProc
End Property
Public Property Let CalleeProc(ByVal val As String)
    pCalleeProc = val
End Property

' Returns a unique identifier for this call-site
Public Function GetID() As String
    GetID = pCallerModule & "." & pCallerProc & "->" & _
            pCalleeModule & "." & pCalleeProc
End Function


