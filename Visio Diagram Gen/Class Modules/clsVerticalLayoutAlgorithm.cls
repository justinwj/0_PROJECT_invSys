VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVerticalLayoutAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsVerticalLayoutAlgorithm
' clsVerticalLayoutAlgorithm (and future ones like clsGridLayoutAlgorithm) handle how to position items.
Option Explicit
'=== Class: clsVerticalLayoutAlgorithm (implements clsLayoutAlgorithm)
Implements clsLayoutAlgorithm

' Default vertical spacing
Private Const DEFAULT_SPACING As Double = 1#

' Simple vertical layout: X=0, Y spaced by DEFAULT_SPACING
Private Function clsLayoutAlgorithm_Layout(ByVal items As Collection, _
                                           ByVal cfg As clsDiagramConfig) As Collection
    Dim out As New Collection
    Dim idx As Long
    Dim itm As clsDiagramItem

    idx = 0
    For Each itm In items
        itm.PosX = cfg.OriginX                         ' Align at configured origin X
        itm.PosY = cfg.OriginY - (idx * cfg.VerticalSpacing) ' Stack downward
        out.Add itm
        idx = idx + 1
    Next itm

    Set clsLayoutAlgorithm_Layout = out
End Function

