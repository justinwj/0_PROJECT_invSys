Attribute VB_Name = "modExportImportAll"
' ===== modExportImportAll.bas =====
'  ExportAllModules
'  ReplaceAllCodeFromFiles
'  ExportTablesHeadersAndControls
'  ExportUserFormControls
'  ImportNewComponentsOnly
Option Explicit
Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fso As Object
    Dim file As Object
    Dim modulesPath As String, classesPath As String, formsPath As String, sheetsPath As String

    exportPath = "D:\justinwj\Workbooks\0_PROJECT_invSys\" ' Change this to your desired folder
    modulesPath = exportPath & "Modules\"
    classesPath = exportPath & "Classes\"
    formsPath = exportPath & "Forms\"
    sheetsPath = exportPath & "Microsoft Excel Objects\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Ensure subfolders exist
    If Not fso.FolderExists(modulesPath) Then fso.CreateFolder modulesPath
    If Not fso.FolderExists(classesPath) Then fso.CreateFolder classesPath
    If Not fso.FolderExists(formsPath) Then fso.CreateFolder formsPath
    If Not fso.FolderExists(sheetsPath) Then fso.CreateFolder sheetsPath

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Module
                vbComp.Export modulesPath & vbComp.Name & ".bas"
            Case 2 ' Class module
                vbComp.Export classesPath & vbComp.Name & ".cls"
            Case 3 ' Form
                vbComp.Export formsPath & vbComp.Name & ".frm"
                ' .frx files are created by Excel, but we will delete them below
            Case 100 ' Microsoft Excel Objects (sheets, workbook)
                vbComp.Export sheetsPath & vbComp.Name & ".cls"
        End Select
    Next vbComp
    ' Delete all .frx files in the export folder and Forms subfolder
    For Each file In fso.GetFolder(formsPath).Files
        If LCase(fso.GetExtensionName(file.Name)) = "frx" Then
            file.Delete True
        End If
    Next file
    MsgBox "Export complete!"
End Sub
' Helper function to check if code block has at least one real code line (not blank, not Attribute/meta)
Private Function HasRealCodeLine(ByVal codeBlock As String) As Boolean
    Dim arr() As String, i As Long, lineTrim As String
    arr = Split(codeBlock, vbCrLf)
    For i = LBound(arr) To UBound(arr)
        lineTrim = Trim(arr(i))
        If lineTrim <> "" Then
            If LCase(Left(lineTrim, 9)) <> "attribute" And _
               UCase(Left(lineTrim, 5)) <> "BEGIN" And _
               UCase(lineTrim) <> "END" And _
               Left(UCase(lineTrim), 7) <> "VERSION" And _
               Left(UCase(lineTrim), 8) <> "MULTIUSE" And _
               Left(lineTrim, 2) <> "//" And _
               (Left(lineTrim, 1) <> "{" Or Right(lineTrim, 1) <> "}") And _
               Left(lineTrim, 7) <> "Caption" And _
               Left(lineTrim, 12) <> "ClientHeight" And _
               Left(lineTrim, 10) <> "ClientLeft" And _
               Left(lineTrim, 9) <> "ClientTop" And _
               Left(lineTrim, 11) <> "ClientWidth" And _
               Left(lineTrim, 13) <> "OleObjectBlob" And _
               Left(lineTrim, 15) <> "StartUpPosition" Then
                HasRealCodeLine = True
                Exit Function
            End If
        End If
    Next i
    HasRealCodeLine = False
End Function
' Replace code in all modules, classes, forms, and sheets from files (no delete/replace)
Sub ReplaceAllCodeFromFiles()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importPath As String
    Dim subFolders(1 To 4) As String
    Dim subFolder As Variant
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
    subFolders(1) = importPath & "Modules\"
    subFolders(2) = importPath & "Classes\"
    subFolders(3) = importPath & "Forms\"
    subFolders(4) = importPath & "Microsoft Excel Objects\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each subFolder In subFolders
        Debug.Print "Scanning folder: " & subFolder
        If Not fso.FolderExists(subFolder) Then
            Debug.Print "Folder does not exist: " & subFolder
            GoTo NextSubFolder
        End If
        Set folder = fso.GetFolder(subFolder)
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
                            If lineTrim Like "Attribute * =*" Or lineTrim Like "Attribute *.* =*" Then GoTo NextLine
                        End If
                        ' Do NOT skip blank lines!
                        ' Do NOT skip lines starting with "#"
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
                    filteredCode = Trim(filteredCode)
                    If HasRealCodeLine(filteredCode) Then
                        With vbComp.CodeModule
                            .DeleteLines 1, .CountOfLines
                            .InsertLines 1, filteredCode
                        End With
                    Else
                        Debug.Print "Skipped InsertLines for " & compName & " (no real code to insert)"
                    End If
                End If
                Set vbComp = Nothing
            End If
        Next file
NextSubFolder:
    Next subFolder
    MsgBox "All code replaced from files!"
End Sub
Sub ExportTablesHeadersAndControls()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim ole As OLEObject
    Dim shp As Shape
    Dim folderPath As String, outputPath As String
    Dim Fnum As Long, hdrs As String
    Dim ctrlType As Long, ctrlTypeName As String
    ' 1) Set your folder (must already exist)
    folderPath = "D:\justinwj\Workbooks\0_PROJECT_invSys\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    ' 2) Append filename
    outputPath = folderPath & "TablesHeadersAndControls.txt"
    Fnum = FreeFile
    Open outputPath For Output As #Fnum
    For Each ws In ThisWorkbook.Worksheets
        Print #Fnum, "Sheet (Tab):  " & ws.Name
        Print #Fnum, "Sheet (Code): " & ws.CodeName
        ' � Tables & Headers �
        For Each lo In ws.ListObjects
            Print #Fnum, "  Table: " & lo.Name
            hdrs = ""
            For Each lc In lo.ListColumns
                hdrs = hdrs & lc.Name & ", "
            Next lc
            If Len(hdrs) > 0 Then hdrs = Left(hdrs, Len(hdrs) - 2)
            Print #Fnum, "    Headers: " & hdrs
        Next lo
        ' � ActiveX Controls �
        For Each ole In ws.OLEObjects
            Print #Fnum, "  ActiveX Control: " & ole.Name & " (" & ole.progID & ")"
            On Error Resume Next
            Print #Fnum, "    LinkedCell: " & ole.LinkedCell
            Print #Fnum, "    TopLeft: " & ole.TopLeftCell.Address(False, False)
            Print #Fnum, "    Caption: " & ole.Object.Caption
            Print #Fnum, "    Value: " & ole.Object.value
            On Error GoTo 0
        Next ole
        ' � Forms Controls �
        For Each shp In ws.Shapes
            If shp.Type = msoFormControl Then
                ctrlType = shp.FormControlType
                Select Case ctrlType
                    Case 0: ctrlTypeName = "Button"
                    Case 1: ctrlTypeName = "Checkbox"
                    Case 2: ctrlTypeName = "DropDown"
                    Case 3: ctrlTypeName = "EditBox"
                    Case 4: ctrlTypeName = "ListBox"
                    Case 5: ctrlTypeName = "ScrollBar"
                    Case 6: ctrlTypeName = "Spinner"
                    Case Else: ctrlTypeName = "Unknown"
                End Select
                Print #Fnum, "  Form Control: " & shp.Name
                Print #Fnum, "    Type: " & ctrlTypeName & " (" & ctrlType & ")"
                On Error Resume Next
                Print #Fnum, "    LinkedCell: " & shp.ControlFormat.LinkedCell
                If shp.HasTextFrame Then
                    Print #Fnum, "    Text: " & Replace(shp.TextFrame.Characters.text, vbCr, " ")
                End If
                On Error GoTo 0
            End If
        Next shp
        Print #Fnum, String(60, "-")
    Next ws
    Close #Fnum
    MsgBox "Export complete:" & vbCrLf & outputPath, vbInformation
End Sub
Sub ExportUserFormControls()
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim ctrl   As MSForms.Control
    Dim outputPath As String, Fnum As Long
    '� adjust folder as needed (must exist) �
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
'   - Scans four organized subfolders (Modules, Classes, Forms,
'     and Microsoft Excel Objects) for any new .bas, .cls, or .frm
'     files.
'   - Imports only those not already in the workbook.
' Configuration: Adjust VSCFolderRoot if your root path changes.
'============================================================
Sub ImportNewComponentsOnly()
    '---- Update this to match your root directory ----
    Const VSCFolderRoot As String = "D:\justinwj\Workbooks\0_PROJECT_invSys\"
    Dim fso        As Object
    Dim vbProj     As VBIDE.VBProject
    Dim subNames   As Variant
    Dim folder     As Object
    Dim file       As Object
    Dim folderPath As String
    Dim compName   As String
    Dim vbComp     As VBIDE.VBComponent
    Dim i          As Long
    ' List of subfolders under the root to scan
    subNames = Array("Modules\", "Classes\", "Forms\", "Microsoft Excel Objects\")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set vbProj = ThisWorkbook.VBProject
    ' Loop through each subfolder
    For i = LBound(subNames) To UBound(subNames)
        folderPath = VSCFolderRoot & subNames(i)
        Debug.Print "Scanning folder: " & folderPath
        If fso.FolderExists(folderPath) Then
            Set folder = fso.GetFolder(folderPath)
            For Each file In folder.Files
                Debug.Print "Found file: " & file.Path
                Select Case LCase(fso.GetExtensionName(file.Name))
                    Case "bas", "cls", "frm"
                        compName = fso.GetBaseName(file.Name)
                        Debug.Print "Processing component: " & compName
                        On Error Resume Next
                        Set vbComp = vbProj.VBComponents(compName)
                        On Error GoTo 0
                        If vbComp Is Nothing Then
                            Debug.Print "Importing component: " & compName
                            vbProj.VBComponents.Import file.Path
                        Else
                            Debug.Print "Component already exists: " & compName
                        End If
                End Select
            Next file
        Else
            Debug.Print "Folder does not exist: " & folderPath
        End If
    Next i
    MsgBox "Import complete. All new components have been added to the workbook.", vbInformation
End Sub

