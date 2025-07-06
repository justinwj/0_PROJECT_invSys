Attribute VB_Name = "modDiagnose"
' modDiagnose -- tools for development
' Diagnostic routine to list all Masters in a Visio stencil and check for a specific master
Option Explicit

Sub Diag_ListStencilMasters(Optional ByVal masterToFind As String = "Container 1")
    Const visBuiltInStencilContainers As Long = 2
    Const visMSUS As Long = 0
    Const visOpenHidden As Long = 64

    Dim visApp As Object, stencilPath As String, stn As Object
    Dim m As Object, found As Boolean

    On Error GoTo ErrHandler

    Set visApp = CreateObject("Visio.Application")
    stencilPath = visApp.GetBuiltInStencilFile(visBuiltInStencilContainers, visMSUS)
    Set stn = visApp.Documents.OpenEx(stencilPath, visOpenHidden)

    Debug.Print "Masters in stencil (" & stencilPath & "):"
    found = False
    For Each m In stn.Masters
        Debug.Print "  - " & m.NameU
        If LCase$(m.NameU) = LCase$(masterToFind) Then found = True
    Next

    If found Then
        Debug.Print "Master '" & masterToFind & "' FOUND in stencil."
    Else
        Debug.Print "Master '" & masterToFind & "' NOT FOUND in stencil!"
        MsgBox "Master '" & masterToFind & "' not found in stencil: " & stencilPath, vbExclamation
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'Call it once from Immediate to see every master in any stencil
' ListStencilMasters Application.GetBuiltInStencilFile(23, 0)   'Basic Flowchart
Sub ListStencilMasters(stencilPath As String)
    Dim stn As Object, m As Object
    Set stn = Application.Documents.OpenEx(stencilPath, 64)   '64 = visOpenHidden
    Debug.Print "Masters in stencil (" & stencilPath & "):"
    For Each m In stn.Masters
        Debug.Print "  - "; m.NameU
    Next m
    stn.Close
End Sub

Sub Diag_ListBasicUMasters()
    Const stencilU$ = "BASIC_U.vssx"
    Const visOpenHidden As Long = 64
    Dim visApp As Object, stn As Object, m As Object

    Set visApp = CreateObject("Visio.Application")
    Set stn = visApp.Documents.OpenEx(stencilU, visOpenHidden)
    Debug.Print "Masters in stencil (" & stencilU & "):"
    For Each m In stn.Masters
        Debug.Print "  - " & m.NameU
    Next
    stn.Close
End Sub

' Diag_ListAllStencilMasters: Enumerates all Visio stencil masters in the Visio Content folder
' and extracts extended metadata to a worksheet for downstream use in modDiagramCore.
Public Sub Diag_ListAllStencilMasters()
    Const visOpenHidden As Long = 64
    Dim fso        As Object     ' Scripting.FileSystemObject
    Dim rootFolder As Object     ' Scripting.Folder
    Dim visApp     As Object     ' Visio.Application
    Dim basePath   As String
    Dim wb         As Workbook
    Dim ws         As Worksheet
    Dim rowIndex   As Long
    Dim calcMode   As XlCalculation
    Dim scUpdt     As Boolean

    ' Performance optimizations
    scUpdt = Application.ScreenUpdating
    Application.ScreenUpdating = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' Prepare output worksheet
    Set wb = ThisWorkbook
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("StencilMasters").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
    ws.Name = "StencilMasters"

    ' Header row
    With ws
        .Cells(1, 1).value = "Stencil File"
        .Cells(1, 2).value = "Master NameU"
        .Cells(1, 3).value = "Master Name"
        .Cells(1, 4).value = "Master ID"
        .Cells(1, 5).value = "Width"
        .Cells(1, 6).value = "Height"
        .Cells(1, 7).value = "Stencil Path"
        .Cells(1, 8).value = "LanguageCode"
    End With
    rowIndex = 2

    ' Determine Visio Content folder (try common paths)
    basePath = Environ$("ProgramFiles") & "\Microsoft Office\root\Office16\Visio Content"
    If Dir(basePath, vbDirectory) = "" Then
        basePath = Environ$("ProgramFiles") & "\Microsoft Office\root\Visio\Visio Content"
    End If
    If Dir(basePath, vbDirectory) = "" Then
        MsgBox "Visio Content folder not found.", vbExclamation
        GoTo CleanUp
    End If

    ' Initialize FileSystemObject and Visio
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(basePath)
    Set visApp = CreateObject("Visio.Application")
    visApp.Visible = False

    ' Extract masters recursively
    ExtractMastersToSheet rootFolder, visApp, visOpenHidden, ws, rowIndex

CleanUp:
    ' Cleanup Visio
    On Error Resume Next
    If Not visApp Is Nothing Then visApp.Quit
    Set visApp = Nothing
    Set rootFolder = Nothing
    Set fso = Nothing

    ' Restore Excel settings
    Application.Calculation = calcMode
    Application.ScreenUpdating = scUpdt

    MsgBox "Completed extracting stencil masters with metadata. See 'StencilMasters' sheet.", vbInformation
End Sub

' Recursive helper to open each stencil and write metadata
Private Sub ExtractMastersToSheet(folder As Object, visApp As Object, openFlags As Long, _
                                   ws As Worksheet, ByRef rowIndex As Long)
    Dim fileItem    As Object    ' Scripting.File
    Dim subFolder   As Object    ' Scripting.Folder
    Dim stencilDoc  As Object    ' Visio.Document
    Dim masterItem  As Object    ' Visio.Master
    Dim ext         As String
    Dim fPath       As String
    Dim fs          As Object
    Dim LangCode    As String
    Dim widthVal    As Double, heightVal As Double

    Set fs = CreateObject("Scripting.FileSystemObject")

    ' Determine language code from folder name
    LangCode = folder.Name

    For Each fileItem In folder.Files
        ext = LCase$(fs.GetExtensionName(fileItem.Name))
        If ext = "vss" Or ext = "vssx" Then
            fPath = fileItem.Path
            On Error Resume Next
            Set stencilDoc = visApp.Documents.OpenEx(fPath, openFlags)
            If Err.Number = 0 Then
                For Each masterItem In stencilDoc.Masters
                    ' Get default dimensions
                    On Error Resume Next
                    widthVal = masterItem.CellsU("Width").ResultIU
                    heightVal = masterItem.CellsU("Height").ResultIU
                    On Error GoTo 0

                    ws.Cells(rowIndex, 1).value = fileItem.Name
                    ws.Cells(rowIndex, 2).value = masterItem.NameU
                    ws.Cells(rowIndex, 3).value = masterItem.Name
                    ws.Cells(rowIndex, 4).value = masterItem.ID
                    ws.Cells(rowIndex, 5).value = widthVal
                    ws.Cells(rowIndex, 6).value = heightVal
                    ws.Cells(rowIndex, 7).value = fPath
                    ws.Cells(rowIndex, 8).value = LangCode
                    rowIndex = rowIndex + 1
                Next masterItem
                stencilDoc.Close
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next fileItem

    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        ExtractMastersToSheet subFolder, visApp, openFlags, ws, rowIndex
    Next subFolder
End Sub
