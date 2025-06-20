Attribute VB_Name = "modTS_Received"

Option Explicit

'==============================================
' Module: modTS_Received (TS Received Processing)
' Purpose: Process ReceivedTally and invSysData_Receiving without generating new REF_NUMBER
'==============================================

Public Sub ProcessReceivedBatch()
    Dim wsRecv    As Worksheet: Set wsRecv    = ThisWorkbook.Sheets("ReceivedTally")
    Dim tblRecv   As ListObject: Set tblRecv   = wsRecv.ListObjects("ReceivedTally")
    Dim tblDet    As ListObject: Set tblDet    = wsRecv.ListObjects("invSysData_Receiving")
    Dim wsInv     As Worksheet: Set wsInv     = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Dim tblInv    As ListObject: Set tblInv    = wsInv.ListObjects("invSys")
    Dim wsLog     As Worksheet: Set wsLog     = ThisWorkbook.Sheets("ReceivedLog")
    Dim tblLog    As ListObject: Set tblLog    = wsLog.ListObjects("ReceivedLog")

    Dim rowCount  As Long:       rowCount     = tblRecv.ListRows.Count
    Dim j         As Long
    Dim refNum    As String
    Dim items     As String
    Dim qty       As Double
    Dim price     As Double
    Dim itemCode  As String
    Dim rowNum    As Long
    Dim uom       As String
    Dim vendor    As String
    Dim location  As String
    Dim entryDate As Date
    Dim newRow    As ListRow

    ' Process each matching row in both staging tables
    For j = 1 To rowCount
        ' 1) Read REF_NUMBER, ITEMS, QUANTITY, PRICE from ReceivedTally
        With tblRecv.DataBodyRange
            refNum = CStr(.Cells(j, tblRecv.ListColumns("REF_NUMBER").Index).Value)
            items  = CStr(.Cells(j, tblRecv.ListColumns("ITEMS").Index).Value)
            qty    = CDbl(.Cells(j, tblRecv.ListColumns("QUANTITY").Index).Value)
            price  = CDbl(.Cells(j, tblRecv.ListColumns("PRICE").Index).Value)
        End With

        ' 2) Read ROW, ITEM_CODE, UOM, VENDOR, LOCATION, ENTRY_DATE from invSysData_Receiving
        With tblDet.DataBodyRange
            rowNum    = CLng(.Cells(j, tblDet.ListColumns("ROW").Index).Value)
            itemCode  = CStr(.Cells(j, tblDet.ListColumns("ITEM_CODE").Index).Value)
            uom       = CStr(.Cells(j, tblDet.ListColumns("UOM").Index).Value)
            vendor    = CStr(.Cells(j, tblDet.ListColumns("VENDOR").Index).Value)
            location  = CStr(.Cells(j, tblDet.ListColumns("LOCATION").Index).Value)
            entryDate = CDate(.Cells(j, tblDet.ListColumns("ENTRY_DATE").Index).Value)
        End With

        ' 3) Append to ReceivedLog using the existing REF_NUMBER
        Set newRow = tblLog.ListRows.Add
        With tblLog.ListColumns
            newRow.Range(1, .Item("REF_NUMBER").Index ).Value = refNum
            newRow.Range(1, .Item("ITEMS").Index      ).Value = items
            newRow.Range(1, .Item("QUANTITY").Index   ).Value = qty
            newRow.Range(1, .Item("PRICE").Index      ).Value = price
            newRow.Range(1, .Item("UOM").Index        ).Value = uom
            newRow.Range(1, .Item("VENDOR").Index     ).Value = vendor
            newRow.Range(1, .Item("LOCATION").Index   ).Value = location
            newRow.Range(1, .Item("ITEM_CODE").Index  ).Value = itemCode
            newRow.Range(1, .Item("ROW").Index        ).Value = rowNum
            newRow.Range(1, .Item("ENTRY_DATE").Index ).Value = entryDate
        End With

        ' 4) Update inventory RECEIVED column in invSys table
        With tblInv.ListRows(rowNum).Range
            .Cells(tblInv.ListColumns("RECEIVED").Index).Value = _
                Val(.Cells(tblInv.ListColumns("RECEIVED").Index).Value) + qty
        End With
    Next j

    ' 5) Clear staging tables
    If Not tblRecv.DataBodyRange Is Nothing Then tblRecv.DataBodyRange.Delete
    If Not tblDet.DataBodyRange Is Nothing Then tblDet.DataBodyRange.Delete
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
