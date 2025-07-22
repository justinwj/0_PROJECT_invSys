VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiagramItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDiagramItem
Option Explicit

' Class module: clsDiagramItem
' Represents a single diagram element for rendering

' Private backing variables
Private pStencilNameU As String
Private pLabelText    As String
Private pMeta As clsMasterMeta
Private pShape As Visio.Shape
Private pPosX          As Double
Private pPosY          As Double

' StencilNameU property
Public Property Get StencilNameU() As String
    StencilNameU = pStencilNameU
End Property
Public Property Let StencilNameU(ByVal Value As String)
    pStencilNameU = Value
End Property

' LabelText property
Public Property Get LabelText() As String
    LabelText = pLabelText
End Property
Public Property Let LabelText(ByVal Value As String)
    pLabelText = Value
End Property

'--- Meta Property ---
Public Property Get meta() As clsMasterMeta
    Set meta = pMeta
End Property

Public Property Set meta(ByVal val As clsMasterMeta)
    Set pMeta = val
End Property

'--- Shape Property (read-only) ---
Public Property Get Shape() As Visio.Shape
    Set Shape = pShape
End Property

Private Property Set Shape(ByVal val As Visio.Shape)
    Set pShape = val
End Property

' PosX property
Public Property Get PosX() As Double
    PosX = pPosX
End Property
Public Property Let PosX(ByVal Value As Double)
    pPosX = Value
End Property

' PosY property
Public Property Get PosY() As Double
    PosY = pPosY
End Property
Public Property Let PosY(ByVal Value As Double)
    pPosY = Value
End Property

'--- Draws this item onto the specified page using the shape provider ---
Public Sub Draw(ByVal pg As Visio.page, ByVal shapeProv As clsShapeProvider)
    Dim m As Visio.master
    Set m = shapeProv.GetMaster(pMeta.Path, pMeta.DisplayNameU)
    Set pShape = pg.Drop(m, pPosX, pPosY)
End Sub
