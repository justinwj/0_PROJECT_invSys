Attribute VB_Name = "modExportImportAll"
' ===== modExportImportAll.bas =====
'  ExportAllModules
'  ReplaceAllCodeFromFiles
'  ExportTablesHeadersAndControls
'  ExportUserFormControls
'  ImportNewComponentsOnly
Option Explicit
Sub ExportAllModules()
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim fso As Object
    Dim fileItem As Object

    ' Root folder path; ensure subfolders "Sheets", "Forms", "Modules", and "Class Modules" exist.
    exportPath = "D:\justinwj\Workbooks\0_PROJECT_invSys"
    ' Ensure trailing backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"

    ' Make sure Excel is set to allow programmatic access to VBProject
    ' (Trust Center > Macro Settings > Trust access to the VBA project object model)

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule ' Standard Modules
                On Error Resume Next
                vbComp.Export exportPath & "Modules\" & vbComp.Name & ".bas"
                On Error GoTo 0

            Case vbext_ct_ClassModule ' Class Modules
                On Error Resume Next
                vbComp.Export exportPath & "Class Modules\" & vbComp.Name & ".cls"
                On Error GoTo 0

            Case vbext_ct_MSForm ' UserForms
                On Error Resume Next
                vbComp.Export exportPath & "Forms\" & vbComp.Name & ".frm"
                On Error GoTo 0

            Case vbext_ct_Document ' Sheets and ThisWorkbook
                On Error Resume Next
                vbComp.Export exportPath & "Sheets\" & vbComp.Name & ".cls"
                On Error GoTo 0
        End Select
    Next vbComp

    ' Remove FRX files from the Forms folder, if present
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(exportPath & "Forms") Then
        For Each fileItem In fso.GetFolder(exportPath & "Forms").Files
            If LCase(fso.GetExtensionName(fileItem.Name)) = "frx" Then
                fileItem.Delete True
            End If
        Next fileItem
    End If

    MsgBox "Export complete!"
End Sub
' Refactored to work with subfolders: Sheets, Forms, Modules, Class Modules
Public Sub ReplaceAllCodeFromFiles()
    Const BASE_PATH As String = "D:\justinwj\Workbooks\0_PROJECT_invSys\"  ' ? adjust as needed
    Dim fso            As Object
    Dim wbProj         As VBIDE.VBProject
    Dim subFolders     As Variant, sf As Variant
    Dim folder         As Object, fileItem As Object
    Dim vbComp         As VBIDE.VBComponent
    Dim ts             As Object
    Dim codeText       As String
    Dim compName       As String
    Dim ext            As String
    
    ' Early exit if base path missing
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right(BASE_PATH, 1) <> "\" Or Not fso.FolderExists(BASE_PATH) Then
        MsgBox "Base folder not found: " & BASE_PATH, vbExclamation
        Exit Sub
    End If
    
    Set wbProj = ThisWorkbook.VBProject
    
    ' 1) Replace Sheet modules (only code inside existing sheet components)
    If fso.FolderExists(BASE_PATH & "Sheets") Then
        Set folder = fso.GetFolder(BASE_PATH & "Sheets")
        For Each fileItem In folder.Files
            If LCase(fso.GetExtensionName(fileItem.Name)) = "cls" Then
                compName = fso.GetBaseName(fileItem.Name)
                On Error Resume Next
                Set vbComp = wbProj.VBComponents(compName)
                On Error GoTo 0
                If Not vbComp Is Nothing Then
                    ' Read the .cls file text
                    Set ts = fso.OpenTextFile(fileItem.Path, 1)
                    codeText = ts.ReadAll
    '--- after you read codeText = ts.ReadAll

' Split into lines
Dim arr() As String, i As Long, metaEnd As Long
arr = Split(codeText, vbCrLf)

' Find where the blank line follows the metadata header
For i = LBound(arr) To UBound(arr)
    If Trim(arr(i)) = "" Then
        metaEnd = i
        Exit For
    End If
Next i

' Rebuild codeText from the line after metadata
If metaEnd > 0 Then
    Dim body As String
    For i = metaEnd + 1 To UBound(arr)
        body = body & arr(i) & vbCrLf
    Next i
    codeText = body
End If

' Now delete the old lines and insert only body
With vbComp.CodeModule
    .DeleteLines 1, .CountOfLines
    .InsertLines 1, codeText
End With
                
                    ts.Close
                    ' Replace all lines in the sheet module
                    With vbComp.CodeModule
                        .DeleteLines 1, .CountOfLines
                        .InsertLines 1, codeText
                    End With
                End If
            End If
        Next fileItem
    End If
    
    ' 2) For Forms, Standard Modules, and Class Modules: remove & re-import
    subFolders = Array( _
        Array("Forms", "frm", "vbext_ct_MSForm"), _
        Array("Modules", "bas", "vbext_ct_StdModule"), _
        Array("Class Modules", "cls", "vbext_ct_ClassModule") _
    )
    For Each sf In subFolders
        If fso.FolderExists(BASE_PATH & sf(0)) Then
            Set folder = fso.GetFolder(BASE_PATH & sf(0))
            For Each fileItem In folder.Files
                ext = LCase(fso.GetExtensionName(fileItem.Name))
                If ext = sf(1) Then
                    compName = fso.GetBaseName(fileItem.Name)
                    ' Remove old component if it exists
                    On Error Resume Next
                    Set vbComp = wbProj.VBComponents(compName)
                    If Not vbComp Is Nothing Then
                        wbProj.VBComponents.Remove vbComp
                    End If
                    On Error GoTo 0
                    ' Import the new component file
                    wbProj.VBComponents.Import fileItem.Path
                End If
            Next fileItem
        End If
    Next sf
    
    MsgBox "All VBA code replaced from exported folders!", vbInformation
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
        ' ? Tables & Headers ?
        For Each lo In ws.ListObjects
            Print #Fnum, "  Table: " & lo.Name
            hdrs = ""
            For Each lc In lo.ListColumns
                hdrs = hdrs & lc.Name & ", "
            Next lc
            If Len(hdrs) > 0 Then hdrs = Left(hdrs, Len(hdrs) - 2)
            Print #Fnum, "    Headers: " & hdrs
        Next lo
        ' ? ActiveX Controls ?
        For Each ole In ws.OLEObjects
            Print #Fnum, "  ActiveX Control: " & ole.Name & " (" & ole.progID & ")"
            On Error Resume Next
            Print #Fnum, "    LinkedCell: " & ole.LinkedCell
            Print #Fnum, "    TopLeft: " & ole.TopLeftCell.Address(False, False)
            Print #Fnum, "    Caption: " & ole.Object.Caption
            Print #Fnum, "    Value: " & ole.Object.value
            On Error GoTo 0
        Next ole
        ' ? Forms Controls ?
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
    '? adjust folder as needed (must exist) ?
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

' Requires reference to “Microsoft Visual Basic for Applications Extensibility 5.3”
' and Trust Center > Macro Settings > “Trust access to the VBA project object model” enabled.

Public Sub ExportAllCodeToSingleFiles()
    Dim exportPath As String
    Dim wsFileNum   As Long, frmFileNum As Long
    Dim clsFileNum  As Long, modFileNum As Long
    Dim vbComp      As VBIDE.VBComponent
    Dim codeMod     As VBIDE.CodeModule
    
    ' ? Modify this to your desired folder (must already exist)
    exportPath = "D:\justinwj\Workbooks\0_PROJECT_invSys"
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Open our four output files
    wsFileNum = FreeFile: Open exportPath & "SheetsCode.txt" For Output As #wsFileNum
    frmFileNum = FreeFile: Open exportPath & "FormsCode.txt" For Output As #frmFileNum
    clsFileNum = FreeFile: Open exportPath & "ClassModulesCode.txt" For Output As #clsFileNum
    modFileNum = FreeFile: Open exportPath & "StandardModulesCode.txt" For Output As #modFileNum
    
    ' Loop through every component in this workbook
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Set codeMod = vbComp.CodeModule
        Select Case vbComp.Type
            Case vbext_ct_Document           ' Sheets & ThisWorkbook
                Print #wsFileNum, "''''''''''''''''''''''''''''''''''''"
                Print #wsFileNum, "' Component: " & vbComp.Name
                Print #wsFileNum, "''''''''''''''''''''''''''''''''''''"
                If codeMod.CountOfLines > 0 Then
                    Print #wsFileNum, codeMod.lines(1, codeMod.CountOfLines)
                End If
                Print #wsFileNum, vbCrLf
            
            Case vbext_ct_MSForm             ' UserForms
                Print #frmFileNum, "''''''''''''''''''''''''''''''''''''"
                Print #frmFileNum, "' UserForm: " & vbComp.Name
                Print #frmFileNum, "''''''''''''''''''''''''''''''''''''"
                If codeMod.CountOfLines > 0 Then
                    Print #frmFileNum, codeMod.lines(1, codeMod.CountOfLines)
                End If
                Print #frmFileNum, vbCrLf
            
            Case vbext_ct_ClassModule        ' Class modules
                Print #clsFileNum, "''''''''''''''''''''''''''''''''''''"
                Print #clsFileNum, "' Class Module: " & vbComp.Name
                Print #clsFileNum, "''''''''''''''''''''''''''''''''''''"
                If codeMod.CountOfLines > 0 Then
                    Print #clsFileNum, codeMod.lines(1, codeMod.CountOfLines)
                End If
                Print #clsFileNum, vbCrLf
            
            Case vbext_ct_StdModule          ' Standard (.bas) modules
                Print #modFileNum, "''''''''''''''''''''''''''''''''''''"
                Print #modFileNum, "' Module: " & vbComp.Name
                Print #modFileNum, "''''''''''''''''''''''''''''''''''''"
                If codeMod.CountOfLines > 0 Then
                    Print #modFileNum, codeMod.lines(1, codeMod.CountOfLines)
                End If
                Print #modFileNum, vbCrLf
        End Select
    Next vbComp
    
    ' Close all files
    Close #wsFileNum
    Close #frmFileNum
    Close #clsFileNum
    Close #modFileNum
    
    MsgBox "All code exported to:" & vbCrLf & _
           exportPath & vbCrLf & _
           "(SheetsCode.txt, FormsCode.txt, ClassModulesCode.txt, StandardModulesCode.txt)", _
           vbInformation
End Sub






