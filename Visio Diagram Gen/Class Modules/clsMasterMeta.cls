VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMasterMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class module: clsMasterMeta
' Represents metadata for a Visio stencil master

' Private backing variables
Private pFileName As String
Private pDisplayNameU As String
Private pDisplayName As String
Private pID As Long
Private pWidth As Double
Private pHeight As Double
Private pPath As String
Private pLangCode As String

' FileName
Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal Value As String)
    pFileName = Value
End Property

' DisplayNameU (unique key)
Public Property Get DisplayNameU() As String
    DisplayNameU = pDisplayNameU
End Property
Public Property Let DisplayNameU(ByVal Value As String)
    pDisplayNameU = Value
End Property

' DisplayName (human-readable name)
Public Property Get DisplayName() As String
    DisplayName = pDisplayName
End Property
Public Property Let DisplayName(ByVal Value As String)
    pDisplayName = Value
End Property

' ID
Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(ByVal Value As Long)
    pID = Value
End Property

' Width
Public Property Get Width() As Double
    Width = pWidth
End Property
Public Property Let Width(ByVal Value As Double)
    pWidth = Value
End Property

' Height
Public Property Get Height() As Double
    Height = pHeight
End Property
Public Property Let Height(ByVal Value As Double)
    pHeight = Value
End Property

' Path
Public Property Get Path() As String
    Path = pPath
End Property
Public Property Let Path(ByVal Value As String)
    pPath = Value
End Property

' LangCode
Public Property Get LangCode() As String
    LangCode = pLangCode
End Property
Public Property Let LangCode(ByVal Value As String)
    pLangCode = Value
End Property
