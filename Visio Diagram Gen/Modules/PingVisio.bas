Attribute VB_Name = "PingVisio"
Sub PingVisio()
    Dim visApp As Object
    Set visApp = CreateObject("Visio.Application")
    MsgBox "Visio version: " & visApp.Version          'should show 16.0.xxxx
    visApp.Quit
End Sub
