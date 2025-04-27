Attribute VB_Name = "modExportImportAll"
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
    importPath = "C:\Temp\VBAExport\" ' <-- update as needed
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
                    lineTrim = Trim(codeLines(i))
                        If lineTrim = "" Then GoTo NextLine
                        If UCase(Left(lineTrim, 5)) = "BEGIN" Then GoTo NextLine
                        If UCase(lineTrim) = "END" Then GoTo NextLine
                        If Left(UCase(lineTrim), 7) = "VERSION" Then GoTo NextLine
                        If Left(UCase(lineTrim), 9) = "ATTRIBUTE" Then GoTo NextLine
                        If Left(UCase(lineTrim), 8) = "MULTIUSE" Then GoTo NextLine
                        If Left(lineTrim, 2) = "//" Then GoTo NextLine
                        If Left(lineTrim, 1) = "{" And Right(lineTrim, 1) = "}" Then GoTo NextLine
                        If Left(lineTrim, 7) = "Caption" Then GoTo NextLine
                        If Left(lineTrim, 12) = "ClientHeight" Then GoTo NextLine
                        If Left(lineTrim, 10) = "ClientLeft" Then GoTo NextLine
                        If Left(lineTrim, 9) = "ClientTop" Then GoTo NextLine
                        If Left(lineTrim, 11) = "ClientWidth" Then GoTo NextLine
                        If Left(lineTrim, 13) = "OleObjectBlob" Then GoTo NextLine
                        If Left(lineTrim, 15) = "StartUpPosition" Then GoTo NextLine
                    End If
                    filteredCode = filteredCode & codeLines(i) & vbCrLf
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

