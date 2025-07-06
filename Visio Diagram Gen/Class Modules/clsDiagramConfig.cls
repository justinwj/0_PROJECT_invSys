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

Private pDiagramType  As String
Private pModuleFilter As String
Private pProcFilter   As String
Private pScaleMode    As String
Private pExportFormat As String

' DiagramType property
Public Property Get DiagramType() As String
    DiagramType = pDiagramType
End Property
Public Property Let DiagramType(ByVal value As String)
    pDiagramType = value
End Property

' ModuleFilter property
Public Property Get moduleFilter() As String
    moduleFilter = pModuleFilter
End Property
Public Property Let moduleFilter(ByVal value As String)
    pModuleFilter = value
End Property

' ProcFilter property
Public Property Get procFilter() As String
    procFilter = pProcFilter
End Property
Public Property Let procFilter(ByVal value As String)
    pProcFilter = value
End Property

' ScaleMode property
Public Property Get ScaleMode() As String
    ScaleMode = pScaleMode
End Property
Public Property Let ScaleMode(ByVal value As String)
    pScaleMode = value
End Property

' ExportFormat property
Public Property Get ExportFormat() As String
    ExportFormat = pExportFormat
End Property
Public Property Let ExportFormat(ByVal value As String)
    pExportFormat = value
End Property
