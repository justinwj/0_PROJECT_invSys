VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiagramConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDiagramConnection
' Represents a connection between two diagram items
Option Explicit

Private pFromID           As String
Private pToID             As String
Private pLabelText        As String
Private pConnectorMasterU As String

' FromID
Public Property Get FromID() As String
    FromID = pFromID
End Property
Public Property Let FromID(ByVal Value As String)
    pFromID = Value
End Property

' ToID
Public Property Get ToID() As String
    ToID = pToID
End Property
Public Property Let ToID(ByVal Value As String)
    pToID = Value
End Property

' LabelText
Public Property Get LabelText() As String
    LabelText = pLabelText
End Property
Public Property Let LabelText(ByVal Value As String)
    pLabelText = Value
End Property

' ConnectorMasterU
Public Property Get ConnectorMasterU() As String
    ConnectorMasterU = pConnectorMasterU
End Property
Public Property Let ConnectorMasterU(ByVal Value As String)
    pConnectorMasterU = Value
End Property

