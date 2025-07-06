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
Public Property Let FileName(ByVal value As String)
    pFileName = value
End Property

' DisplayNameU (unique key)
Public Property Get DisplayNameU() As String
    DisplayNameU = pDisplayNameU
End Property
Public Property Let DisplayNameU(ByVal value As String)
    pDisplayNameU = value
End Property

' DisplayName (human-readable name)
Public Property Get DisplayName() As String
    DisplayName = pDisplayName
End Property
Public Property Let DisplayName(ByVal value As String)
    pDisplayName = value
End Property

' ID
Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(ByVal value As Long)
    pID = value
End Property

' Width
Public Property Get Width() As Double
    Width = pWidth
End Property
Public Property Let Width(ByVal value As Double)
    pWidth = value
End Property

' Height
Public Property Get Height() As Double
    Height = pHeight
End Property
Public Property Let Height(ByVal value As Double)
    pHeight = value
End Property

' Path
Public Property Get Path() As String
    Path = pPath
End Property
Public Property Let Path(ByVal value As String)
    pPath = value
End Property

' LangCode
Public Property Get LangCode() As String
    LangCode = pLangCode
End Property
Public Property Let LangCode(ByVal value As String)
    pLangCode = value
End Property
