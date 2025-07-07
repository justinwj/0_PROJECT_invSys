Attribute VB_Name = "modReferences"
Option Explicit

'-------------------------------------------------------------------------------
' Module: modReferences
' Purpose: Programmatically add the VBA project references needed by the diagram generator
' Requirements:
'  - "Trust access to the VBA project object model" must be enabled in Excel Trust Center
'  - Microsoft Visual Basic for Applications Extensibility 5.3 reference must be set (manually once)
'-------------------------------------------------------------------------------
Public Sub AddRequiredReferences()
    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisWorkbook.VBProject
    
    On Error Resume Next
    ' Microsoft Scripting Runtime
    vbProj.References.AddFromGuid _
        GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", _
        Major:=1, Minor:=0
    ' Microsoft Forms 2.0 Object Library
    vbProj.References.AddFromGuid _
        GUID:="{0D452EE1-E08F-101A-852E-02608C4D0BB4}", _
        Major:=2, Minor:=0
    On Error GoTo 0

    ' You can add additional references here as needed
End Sub

