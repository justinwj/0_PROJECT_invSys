Attribute VB_Name = "modTS_Common"
Option Explicit
' This module contains common functions used across the application
Public Function GetUOMFromDataTable(item As String, ItemCode As String, rowNum As String) As String
    Dim ws  As Worksheet
    Dim tbl As ListObject
    Dim findCol As Long
    Dim cel As Range
    
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set tbl = ws.ListObjects("invSysData_Receiving")
    findCol = tbl.ListColumns("ROW").Index
    
    ' Match by ROW
    If rowNum <> "" Then
        For Each cel In tbl.DataBodyRange.Columns(findCol).Cells
            If CStr(cel.Value) = rowNum Then
                GetUOMFromDataTable = cel.Offset(0, tbl.ListColumns("UOM").Index - findCol).Value
                Exit Function
            End If
        Next
    End If
    
    ' Match by ITEM_CODE
    findCol = tbl.ListColumns("ITEM_CODE").Index
    If ItemCode <> "" Then
        For Each cel In tbl.DataBodyRange.Columns(findCol).Cells
            If CStr(cel.Value) = ItemCode Then
                GetUOMFromDataTable = cel.Offset(0, tbl.ListColumns("UOM").Index - findCol).Value
                Exit Function
            End If
        Next
    End If
    
    ' Match by ITEM
    findCol = tbl.ListColumns("ITEM").Index
    For Each cel In tbl.DataBodyRange.Columns(findCol).Cells
        If CStr(cel.Value) = item Then
            GetUOMFromDataTable = cel.Offset(0, tbl.ListColumns("UOM").Index - findCol).Value
            Exit Function
        End If
    Next
    
    GetUOMFromDataTable = ""
End Function

Public Function GetUOMFromInvSys(item As String, ItemCode As String, rowNum As String) As String
    Dim ws  As Worksheet
    Dim tbl As ListObject
    Dim findCol As Long
    Dim cel As Range
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    findCol = tbl.ListColumns("ROW").Index
    If rowNum <> "" Then
        For Each cel In tbl.DataBodyRange.Columns(findCol).Cells
            If CStr(cel.Value) = rowNum Then
                GetUOMFromInvSys = cel.Offset(0, tbl.ListColumns("UOM").Index - findCol).Value
                Exit Function
            End If
        Next
    End If
    
    findCol = tbl.ListColumns("ITEM_CODE").Index
    If ItemCode <> "" Then
        For Each cel In tbl.DataBodyRange.Columns(findCol).Cells
            If CStr(cel.Value) = ItemCode Then
                GetUOMFromInvSys = cel.Offset(0, tbl.ListColumns("UOM").Index - findCol).Value
                Exit Function
            End If
        Next
    End If
    
    findCol = tbl.ListColumns("ITEM").Index
    For Each cel In tbl.DataBodyRange.Columns(findCol).Cells
        If CStr(cel.Value) = item Then
            GetUOMFromInvSys = cel.Offset(0, tbl.ListColumns("UOM").Index - findCol).Value
            Exit Function
        End If
    Next
    
    GetUOMFromInvSys = ""
End Function

