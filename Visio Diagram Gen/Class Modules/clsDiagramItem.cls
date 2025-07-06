VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiagramItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class module: clsDiagramItem
' Represents a single diagram element for rendering

' Private backing variables
Private pStencilNameU As String
Private pLabelText As String
Private pPosX As Double
Private pPosY As Double

' StencilNameU
Public Property Get StencilNameU() As String
    StencilNameU = pStencilNameU
End Property
Public Property Let StencilNameU(ByVal value As String)
    pStencilNameU = value
End Property

' LabelText
Public Property Get LabelText() As String
    LabelText = pLabelText
End Property
Public Property Let LabelText(ByVal value As String)
    pLabelText = value
End Property

' PosX
Public Property Get PosX() As Double
    PosX = pPosX
End Property
Public Property Let PosX(ByVal value As Double)
    pPosX = value
End Property

' PosY
Public Property Get PosY() As Double
    PosY = pPosY
End Property
Public Property Let PosY(ByVal value As Double)
    pPosY = value
End Property

