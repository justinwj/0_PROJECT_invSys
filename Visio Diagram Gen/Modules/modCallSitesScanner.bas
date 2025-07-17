Attribute VB_Name = "modCallSitesScanner"
' modCallSiteScanner.bas
Option Explicit

' Scans a VBComponent for call statements and returns CallSite objects
Public Function ScanModuleForCalls(component As VBIDE.VBComponent) As Collection
    Dim calls As New Collection
    Dim codeMod As VBIDE.CodeModule
    Dim totalLines As Long, lineNum As Long
    Dim regex As Object, matches As Object, m As Object
    Dim procName As String
    Dim cs As clsCallSite

    Set codeMod = component.CodeModule
    totalLines = codeMod.CountOfLines

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b(\w+)\s*\.\s*(\w+)\s*\("  ' matches Module.Proc(
    regex.Global = True

    For lineNum = 1 To totalLines
        Dim lineText As String
        lineText = codeMod.lines(lineNum, 1)
        If regex.Test(lineText) Then
            Set matches = regex.Execute(lineText)
            procName = codeMod.ProcOfLine(lineNum, vbext_pk_Proc)
            For Each m In matches
                Set cs = New clsCallSite
                cs.CallerModule = component.Name
                cs.CallerProc = procName
                cs.CalleeModule = m.SubMatches(0)
                cs.CalleeProc = m.SubMatches(1)
                calls.Add cs
            Next m
        End If
    Next lineNum

    Set ScanModuleForCalls = calls
End Function

' Aggregates call-sites from all modules in the project
Public Function LoadAllCallSites() As Collection
    Dim allCalls As New Collection
    Dim proj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim compCalls As Collection
    Dim cs As clsCallSite

    Set proj = ThisWorkbook.VBProject
    For Each comp In proj.VBComponents
        If comp.Type = vbext_ct_StdModule Or comp.Type = vbext_ct_ClassModule Then
            Set compCalls = ScanModuleForCalls(comp)
            For Each cs In compCalls
                allCalls.Add cs
            Next cs
        End If
    Next comp

    Set LoadAllCallSites = allCalls
End Function
