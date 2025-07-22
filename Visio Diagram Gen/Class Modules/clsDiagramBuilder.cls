VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiagramBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visio constants required for DrawConnections
Const visSectionObject As Integer = 1
Const visRowXForm1      As Integer = 1
Const visRowXForm2      As Integer = 2
Const visXFormPinX      As Integer = 0
Const visXFormPinY      As Integer = 1

Private visApp As Visio.application
Private pg     As Visio.page

Public Sub Initialize(ByVal application As Visio.application, ByVal page As Visio.page)
    Set visApp = application
    Set pg = page
End Sub

Public Function DrawItems(ByVal items As Collection, ByVal shapeProv As clsShapeProvider) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary, itm As clsDiagramItem
    For Each itm In items
        itm.Draw pg, shapeProv
        dict.Add itm.meta.DisplayNameU, itm.Shape
    Next
    Set DrawItems = dict
End Function

Public Sub DrawConnections(ByVal connections As Collection, ByVal shapedDict As Scripting.Dictionary)
    Dim connPair As Variant, fromShp As Visio.Shape, toShp As Visio.Shape, conn As Visio.Shape
    For Each connPair In connections
        Set fromShp = shapedDict(connPair(0))
        Set toShp = shapedDict(connPair(1))
        Set conn = pg.Drop(visApp.ConnectorToolDataObject, 0, 0)
        With conn
            .CellsSRC(visSectionObject, visRowXForm1, visXFormPinX).GlueTo fromShp.CellsSRC(visSectionObject, visRowXForm1, visXFormPinX)
            .CellsSRC(visSectionObject, visRowXForm1, visXFormPinY).GlueTo fromShp.CellsSRC(visSectionObject, visRowXForm1, visXFormPinY)
            .CellsSRC(visSectionObject, visRowXForm2, visXFormPinX).GlueTo toShp.CellsSRC(visSectionObject, visRowXForm1, visXFormPinX)
            .CellsSRC(visSectionObject, visRowXForm2, visXFormPinY).GlueTo toShp.CellsSRC(visSectionObject, visRowXForm1, visXFormPinY)
        End With
    Next
End Sub
