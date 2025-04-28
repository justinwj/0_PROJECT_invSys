Attribute VB_Name = "modExportImportAll"
' ===== modExportImportAll.bas =====
'  ExportAllModules
'  ReplaceAllCodeFromFiles
'  ExportTablesHeadersAndControls
'  ExportUserFormControls
'  ImportNewComponentsOnly
Option Explicit
' Subroutine to export all modules, classes, forms, and Excel objects (sheets, workbook)
Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fso As Object
    Dim file As Object
    exportPath = "D:\justinwj\Workbooks\0_PROJECT_invSys\" ' Change this to your desired folder
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 2 ' Class module
                vbComp.Export exportPath & vbComp.Name & ".cls"
            Case 3 ' Form
                vbComp.Export exportPath & vbComp.Name & ".frm"
                ' .frx files are created by Excel, but we will delete them below
            Case 100 ' Microsoft Excel Objects (sheets, workbook)
                vbComp.Export exportPath & vbComp.Name & ".cls"
        End Select
    Next vbComp
    ' Delete all .frx files in the export folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each file In fso.GetFolder(exportPath).Files
        If LCase(fso.GetExtensionName(file.Name)) = "frx" Then
            file.Delete True
        End If
    Next file
    MsgBox "Export complete!"
End Sub
' Replace code in all modules, classes, forms, and sheets from files (no delete/replace)
Sub ReplaceAllCodeFromFiles()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importPath As String
    Dim vbComp As Object
    Dim compName As String
    Dim ext As String
    Dim codeText As String
    Dim ts As Object
    Dim codeLines() As String
    Dim filteredCode As String
    Dim i As Long
    Dim lineTrim As String
    importPath = "D:\justinwj\Workbooks\0_PROJECT_invSys\" ' <-- update as needed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(importPath) Then
        MsgBox "Folder not found: " & importPath, vbExclamation
        Exit Sub
    End If
    Set folder = fso.GetFolder(importPath)
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            compName = fso.GetBaseName(file.Name)
            ' Find the component by name
            On Error Resume Next
            Set vbComp = ThisWorkbook.VBProject.VBComponents(compName)
            On Error GoTo 0
            If Not vbComp Is Nothing Then
                ' Read file text
                Set ts = fso.OpenTextFile(file.Path, 1)
                codeText = ts.ReadAll
                ts.Close
                codeLines = Split(codeText, vbCrLf)
                filteredCode = ""
                For i = LBound(codeLines) To UBound(codeLines)
                    lineTrim = codeLines(i)
                    ' Filter out meta lines: Attribute lines that look like meta, not code lines
                    If LCase(Left(Trim(lineTrim), 9)) = "attribute" Then
                        ' Only skip if it matches the meta pattern: Attribute <name> = ... or Attribute <name>.<property> = ...
                        If lineTrim Like "Attribute * =*" Or lineTrim Like "Attribute *.* =*" Then GoTo NextLine
                    End If
                    If Trim(lineTrim) = "" Then GoTo NextLine
                    If UCase(Left(Trim(lineTrim), 5)) = "BEGIN" Then GoTo NextLine
                    If UCase(Trim(lineTrim)) = "END" Then GoTo NextLine
                    If Left(UCase(Trim(lineTrim)), 7) = "VERSION" Then GoTo NextLine
                    If Left(UCase(Trim(lineTrim)), 8) = "MULTIUSE" Then GoTo NextLine
                    If Left(Trim(lineTrim), 2) = "//" Then GoTo NextLine
                    If Left(Trim(lineTrim), 1) = "{" And Right(Trim(lineTrim), 1) = "}" Then GoTo NextLine
                    If Left(Trim(lineTrim), 7) = "Caption" Then GoTo NextLine
                    If Left(Trim(lineTrim), 12) = "ClientHeight" Then GoTo NextLine
                    If Left(Trim(lineTrim), 10) = "ClientLeft" Then GoTo NextLine
                    If Left(Trim(lineTrim), 9) = "ClientTop" Then GoTo NextLine
                    If Left(Trim(lineTrim), 11) = "ClientWidth" Then GoTo NextLine
                    If Left(Trim(lineTrim), 13) = "OleObjectBlob" Then GoTo NextLine
                    If Left(Trim(lineTrim), 15) = "StartUpPosition" Then GoTo NextLine
                    filteredCode = filteredCode & lineTrim & vbCrLf
NextLine:
                Next i
                ' Replace code
                With vbComp.CodeModule
                    .DeleteLines 1, .CountOfLines
                    .InsertLines 1, filteredCode
                End With
            End If
            Set vbComp = Nothing
        End If
    Next file
    MsgBox "All code replaced from files!"
End Sub
' exports all tab names and code names for sheet and exports all table names and headers
Sub ExportTablesAndHeaders()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim folderPath As String, outputPath As String
    Dim Fnum As Long, hdrs As String

    ' 1) Set your folder (must already exist)
    folderPath = "D:\justinwj\Workbooks\0_PROJECT_invSys\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    ' 2) Append filename
    outputPath = folderPath & "TablesAndHeaders.txt"

    Fnum = FreeFile
    Open outputPath For Output As #Fnum

    For Each ws In ThisWorkbook.Worksheets
        Print #Fnum, "Sheet (Tab):  " & ws.Name
        Print #Fnum, "Sheet (Code): " & ws.CodeName
        For Each lo In ws.ListObjects
            Print #Fnum, "  Table: " & lo.Name
            hdrs = ""
            For Each lc In lo.ListColumns
                hdrs = hdrs & lc.Name & ", "
            Next lc
            If Len(hdrs) > 0 Then hdrs = Left(hdrs, Len(hdrs) - 2)
            Print #Fnum, "    Headers: " & hdrs
        Next lo
        Print #Fnum, String(40, "-")
    Next ws

    Close #Fnum
    MsgBox "Export complete:" & vbCrLf & outputPath, vbInformation
End Sub

    Sub ExportUserFormControls()
        Dim vbProj As VBIDE.VBProject
        Dim vbComp As VBIDE.VBComponent
        Dim ctrl   As MSForms.Control
        Dim outputPath As String, Fnum As Long
        
        '— adjust folder as needed (must exist) —
        outputPath = "D:\justinwj\Workbooks\0_PROJECT_invSys\UserFormControls.txt"
        Fnum = FreeFile
        Open outputPath For Output As #Fnum
        
        Set vbProj = ThisWorkbook.VBProject
        For Each vbComp In vbProj.VBComponents
            ' only UserForm components
            If vbComp.Type = vbext_ct_MSForm Then
                Print #Fnum, "UserForm: " & vbComp.Name
                ' iterate its controls
                For Each ctrl In vbComp.Designer.Controls
                    Print #Fnum, "  Control: " & ctrl.Name & " (" & TypeName(ctrl) & ")"
                    On Error Resume Next
                    ' many controls have a Caption
                    Print #Fnum, "    Caption: " & ctrl.Caption
                    ' and many have a Value
                    Print #Fnum, "    Value: " & ctrl.value
                    On Error GoTo 0
                Next ctrl
                Print #Fnum, String(50, "-")
            End If
        Next vbComp
        
        Close #Fnum
        MsgBox "UserForm controls exported to:" & vbCrLf & outputPath, vbInformation
    End Sub
    
'============================================================
' Sub: ImportNewComponentsOnly
' Purpose:
'   - Imports any new .bas, .cls, or .frm files from the VSC export
'     root folder into the VBA project.
'   - Moves them into Modules, Classes, Forms, or Sheets subfolders
'     based on their actual component type or PredeclaredId flag.
' Configuration: Update the folder paths as needed.
'============================================================
Sub ImportNewComponentsOnly()
    '---- Update these to match your directory structure ----
    Const VSCFolderRoot   As String = "D:\\justinwj\\Workbooks\\0_PROJECT_invSys\\"
    Const ModulesFolder   As String = "Modules\\"
    Const ClassesFolder   As String = "Classes\\"
    Const FormsFolder     As String = "Forms\\"
    Const SheetsFolder    As String = "Microsoft Excel Objects\\"

    Dim fso      As Object
    Dim vscFld   As Object
    Dim file     As Object
    Dim compName As String
    Dim extType  As String
    Dim vbProj   As VBIDE.VBProject
    Dim vbComp   As VBIDE.VBComponent
    Dim movePath As String
    Dim compType As Long

    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbProj = ThisWorkbook.VBProject
    Set vscFld = fso.GetFolder(VSCFolderRoot)

    ' Loop through files in root folder
    For Each file In vscFld.Files
        extType = LCase(fso.GetExtensionName(file.Name))
        If extType = "bas" Or extType = "cls" Or extType = "frm" Then
            compName = fso.GetBaseName(file.Name)
            ' Skip if component already exists
            On Error Resume Next
            Set vbComp = vbProj.VBComponents(compName)
            On Error GoTo 0

            If vbComp Is Nothing Then
                ' Import the component
                On Error Resume Next
                vbProj.VBComponents.Import file.Path
                If Err.Number <> 0 Then
                    Debug.Print "Failed import: " & file.Name & " => " & Err.Description
                    Err.Clear
                    GoTo NextFile
                End If
                On Error GoTo 0

                ' Retrieve newly imported component
                Set vbComp = vbProj.VBComponents(compName)

                ' Determine effective type: treat class modules with PredeclaredId=True as sheets
                compType = vbComp.Type
                On Error Resume Next
                If compType = vbext_ct_ClassModule Then
                    If vbComp.Properties("VB_PredeclaredId") = True Then
                        compType = vbext_ct_Document
                    End If
                End If
                On Error GoTo 0

                ' Choose move path based on compType
                Select Case compType
                    Case vbext_ct_StdModule   ' .bas
                        movePath = VSCFolderRoot & ModulesFolder
                    Case vbext_ct_ClassModule ' .cls class modules
                        movePath = VSCFolderRoot & ClassesFolder
                    Case vbext_ct_MSForm      ' .frm
                        movePath = VSCFolderRoot & FormsFolder
                    Case vbext_ct_Document    ' sheets (incl. Predeclared class modules)
                        movePath = VSCFolderRoot & SheetsFolder
                    Case Else
                        movePath = VSCFolderRoot
                End Select
                ' Ensure trailing backslash
                If Right(movePath, 1) <> "\\" Then movePath = movePath & "\\"

                ' Move the file
                On Error Resume Next
                fso.MoveFile file.Path, movePath & file.Name
                On Error GoTo 0
            End If
        End If
NextFile:
    Next file

    MsgBox "Import complete. New components added and organized.", vbInformation
End Sub


