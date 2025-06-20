Attribute VB_Name = "modTS_Received"

Option Explicit

'────────────────────────────────────────────────────────────
' Main routine: log, update inventory, then clear staging
'────────────────────────────────────────────────────────────
Public Sub ProcessReceivedBatch()
    Dim wsRecv        As Worksheet:   Set wsRecv  = ThisWorkbook.Sheets("ReceivedTally")
    Dim tblRecv       As ListObject:  Set tblRecv = wsRecv.ListObjects("ReceivedTally")
    Dim tblDetail     As ListObject:  Set tblDetail = wsRecv.ListObjects("invSysData_Receiving")
    Dim wsInv         As Worksheet:   Set wsInv   = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Dim tblInv        As ListObject:  Set tblInv  = wsInv.ListObjects("invSys")
    Dim wsLog         As Worksheet:   Set wsLog   = ThisWorkbook.Sheets("ReceivedLog")
    Dim tblLog        As ListObject:  Set tblLog  = wsLog.ListObjects("ReceivedLog")
    Dim batchRef      As String:      batchRef   = modTS_Log.GenerateOrderNumber()
    
    Dim lrRecv    As ListRow
    Dim lrDet     As ListRow
    Dim newRow    As ListRow
    Dim qty       As Double
    Dim price     As Double
    Dim rowNum    As Long
    Dim uom       As String
    Dim vendor    As String
    Dim location  As String
    Dim entryDate As Date
    
    Debug.Print "ProcessReceivedBatch: found " & tblRecv.ListRows.Count & " staging rows."
    
    For Each lrRecv In tblRecv.ListRows
        ' read key fields from ReceivedTally
        rowNum = CLng(lrRecv.Range.Cells(1, tblRecv.ListColumns("ROW").Index).Value)
        qty    = CDbl(lrRecv.Range.Cells(1, tblRecv.ListColumns("QUANTITY").Index).Value)
        price  = CDbl(lrRecv.Range.Cells(1, tblRecv.ListColumns("PRICE").Index).Value)
        
        ' pull the rest from invSysData_Receiving
        uom       = ""
        vendor    = ""
        location  = ""
        entryDate = Now
        For Each lrDet In tblDetail.ListRows
            If CLng(lrDet.Range.Cells(1, tblDetail.ListColumns("ROW").Index).Value) = rowNum Then
                With lrDet.Range
                    uom       = CStr(.Cells(tblDetail.ListColumns("UOM").Index).Value)
                    vendor    = CStr(.Cells(tblDetail.ListColumns("VENDOR").Index).Value)
                    location  = CStr(.Cells(tblDetail.ListColumns("LOCATION").Index).Value)
                    entryDate = CDate(.Cells(tblDetail.ListColumns("ENTRY_DATE").Index).Value)
                End With
                Exit For
            End If
        Next lrDet
        
        ' 1) Append to ReceivedLog
        Set newRow = tblLog.ListRows.Add
        With tblLog.ListColumns
            newRow.Range(1, .Item("REF_NUMBER").Index ).Value = batchRef
            newRow.Range(1, .Item("ITEMS").Index      ).Value = CStr(lrRecv.Range.Cells(1, tblRecv.ListColumns("ITEMS").Index).Value)
            newRow.Range(1, .Item("QUANTITY").Index   ).Value = qty
            newRow.Range(1, .Item("PRICE").Index      ).Value = price
            newRow.Range(1, .Item("UOM").Index        ).Value = uom
            newRow.Range(1, .Item("VENDOR").Index     ).Value = vendor
            newRow.Range(1, .Item("LOCATION").Index   ).Value = location
            newRow.Range(1, .Item("ITEM_CODE").Index  ).Value = CStr(lrRecv.Range.Cells(1, tblRecv.ListColumns("ITEM_CODE").Index).Value)
            newRow.Range(1, .Item("ROW").Index        ).Value = rowNum
            newRow.Range(1, .Item("ENTRY_DATE").Index ).Value = entryDate
        End With
        
        ' 2) Update INVENTORY MANAGEMENT → RECEIVED
        With tblInv.ListRows(rowNum).Range
            .Cells(tblInv.ListColumns("RECEIVED").Index).Value = _
                Val(.Cells(tblInv.ListColumns("RECEIVED").Index).Value) + qty
        End With
        
        Debug.Print " → Row " & rowNum & ": qty=" & qty & " logged & added."
    Next lrRecv
    
    ' 3) clear staging
    If Not tblRecv.DataBodyRange Is Nothing Then tblRecv.DataBodyRange.Delete
    If Not tblDetail.DataBodyRange Is Nothing Then tblDetail.DataBodyRange.Delete
    
    Debug.Print "ProcessReceivedBatch: done. Staging cleared."
End Sub

'────────────────────────────────────────────────────────────
' Pulls UOM, VENDOR, LOCATION, ENTRY_DATE from invSysData_Receiving
'────────────────────────────────────────────────────────────
Public Sub GetReceivingDetails( _
    ByVal itemCode  As String, _
    ByVal rowNum    As Long, _
    ByRef uom       As String, _
    ByRef vendor    As String, _
    ByRef location  As String, _
    ByRef entryDate As Date)

    Dim wsTable As Worksheet
    Dim tbl      As ListObject
    Dim lr       As ListRow
    Dim colUOM       As Long, colVendor As Long
    Dim colLocation  As Long, colRow As Long

    Set wsTable = ThisWorkbook.Sheets("ReceivedTally")
    Set tbl     = wsTable.ListObjects("invSysData_Receiving")

    ' Find column indexes once
    colUOM      = tbl.ListColumns("UOM").Index
    colVendor   = tbl.ListColumns("VENDOR").Index
    colLocation = tbl.ListColumns("LOCATION").Index
    colRow      = tbl.ListColumns("ROW").Index

    ' Default fallback
    uom       = ""
    vendor    = ""
    location  = ""
    entryDate = Now

    For Each lr In tbl.ListRows
        With lr.Range
            If .Cells(colRow).Value = rowNum Then
                uom       = CStr(.Cells(colUOM     ).Value)
                vendor    = CStr(.Cells(colVendor  ).Value)
                location  = CStr(.Cells(colLocation).Value)
                entryDate = CDate(.Cells(tbl.ListColumns("ENTRY_DATE").Index).Value)
                Exit Sub
            End If
        End With
    Next lr
End Sub


'────────────────────────────────────────────────────────────
' Appends a single row into the ReceivedLog table
'────────────────────────────────────────────────────────────
Public Sub AppendReceivedLogRecord( _
    ByVal refNum    As String, _
    ByVal itemName  As String, _
    ByVal qty       As Double, _
    ByVal price     As Double, _
    ByVal uom       As String, _
    ByVal vendor    As String, _
    ByVal location  As String, _
    ByVal itemCode  As String, _
    ByVal rowNum    As Long, _
    ByVal entryDate As Date)

    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow

    Set wsLog  = ThisWorkbook.Sheets("ReceivedLog")
    Set tblLog = wsLog.ListObjects("ReceivedLog")

    ' Debug to confirm we’re appending to the right table
    Debug.Print "[AppendReceivedLogRecord] sheet=" & wsLog.Name & "; table=" & tblLog.Name

    Set newRow = tblLog.ListRows.Add
    With tblLog.ListColumns
        newRow.Range(1, .Item("REF_NUMBER").Index ).Value = refNum
        newRow.Range(1, .Item("ITEMS").Index      ).Value = itemName
        newRow.Range(1, .Item("QUANTITY").Index   ).Value = qty
        newRow.Range(1, .Item("PRICE").Index      ).Value = price
        newRow.Range(1, .Item("UOM").Index        ).Value = uom
        newRow.Range(1, .Item("VENDOR").Index     ).Value = vendor
        newRow.Range(1, .Item("LOCATION").Index   ).Value = location
        newRow.Range(1, .Item("ITEM_CODE").Index  ).Value = itemCode
        newRow.Range(1, .Item("ROW").Index        ).Value = rowNum
        newRow.Range(1, .Item("ENTRY_DATE").Index ).Value = entryDate
    End With
End Sub
