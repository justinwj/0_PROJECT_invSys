VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiagramConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDiagramConfig
' Class data for modDiagramCore
' Class Module: clsDiagramConfig
' Holds configuration parameters for diagram generation
Option Explicit

' Private backing fields
Private pDiagramType   As String
Private pModuleFilter  As String
Private pProcFilter    As String
Private pScaleMode     As String
Private pExportFormat  As String

' Initialize default values when class is created
Private Sub Class_Initialize()
    pDiagramType = "CallGraph"
    pModuleFilter = "*"
    pProcFilter = "*"
    pScaleMode = "FitToPage"
    pExportFormat = "PNG"
End Sub

' Property Get/Let for DiagramType
Public Property Get DiagramType() As String
    DiagramType = pDiagramType
End Property
Public Property Let DiagramType(ByVal Value As String)
    pDiagramType = Value
End Property

' Property Get/Let for ModuleFilter
Public Property Get moduleFilter() As String
    moduleFilter = pModuleFilter
End Property
Public Property Let moduleFilter(ByVal Value As String)
    pModuleFilter = Value
End Property

' Property Get/Let for ProcFilter
Public Property Get procFilter() As String
    procFilter = pProcFilter
End Property
Public Property Let procFilter(ByVal Value As String)
    pProcFilter = Value
End Property

' Property Get/Let for ScaleMode
Public Property Get ScaleMode() As String
    ScaleMode = pScaleMode
End Property
Public Property Let ScaleMode(ByVal Value As String)
    pScaleMode = Value
End Property

' Property Get/Let for ExportFormat
Public Property Get ExportFormat() As String
    ExportFormat = pExportFormat
End Property
Public Property Let ExportFormat(ByVal Value As String)
    pExportFormat = Value
End Property
