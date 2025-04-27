Attribute VB_Name = "modTS_Log"
' ==============================================
' Module: modTS_Log (TS stands for Tally System)
' ==============================================
Option Explicit
' Shared function for generating unique order numbers
    Function GenerateOrderNumber() As String
        GenerateOrderNumber = "ORD" & Format(Now(), "YYMMDDHHMMSS")
    End Function
Sub LogShipments(shipmentSummary As Object)
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In shipmentSummary.Keys
        Set newRow = ThisWorkbook.Sheets("ShipmentsLog").ListObjects("ShipmentsLog").ListRows.Add
        newRow.Range(1, 1).value = GenerateOrderNumber()
        newRow.Range(1, 2).value = key
        newRow.Range(1, 3).value = shipmentSummary(key)
        newRow.Range(1, 4).value = Now()
    Next key
End Sub
Sub LogReceived(receivedSummary As Object)
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In receivedSummary.Keys
        Set newRow = ThisWorkbook.Sheets("ReceivedLog").ListObjects("ReceivedLog").ListRows.Add
        newRow.Range(1, 1).value = GenerateOrderNumber()
        newRow.Range(1, 2).value = key
        newRow.Range(1, 3).value = receivedSummary(key)
        newRow.Range(1, 4).value = Now()
    Next key
End Sub
Sub LogReceivedDetailed(receivedSummary As Object)
    On Error GoTo ErrorHandler
    Dim key As Variant
    Dim newRow As ListRow
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("ReceivedLog")
    Set tbl = ws.ListObjects("ReceivedLog")
    ' Check that the table exists
    If tbl Is Nothing Then
        MsgBox "ReceivedLog table not found!", vbExclamation
        Exit Sub
    End If
    ' Get column indexes
    Dim colRefNum As Long, colItems As Long, colQty As Long, colPrice As Long
    Dim colUOM As Long, colVendor As Long, colLocation As Long
    Dim colItemCode As Long, colRow As Long, colEntryDate As Long
    For colRefNum = 1 To tbl.ListColumns.count
        Select Case UCase(tbl.ListColumns(colRefNum).Name)
            Case "REF_NUMBER": colRefNum = colRefNum
            Case "ITEMS": colItems = colRefNum
            Case "QUANTITY": colQty = colRefNum
            Case "PRICE": colPrice = colRefNum
            Case "UOM": colUOM = colRefNum
            Case "VENDOR": colVendor = colRefNum
            Case "LOCATION": colLocation = colRefNum
            Case "ITEM_CODE": colItemCode = colRefNum
            Case "ROW": colRow = colRefNum
            Case "ENTRY_DATE": colEntryDate = colRefNum
        End Select
    Next
    ' Log each item to the ReceivedLog table
    Application.ScreenUpdating = False
    For Each key In receivedSummary.Keys
        Dim itemData As Variant
        itemData = receivedSummary(key)
        ' Add new row to the table
        Set newRow = tbl.ListRows.Add
        ' Fill in all fields - Make sure array indexes match your data structure
        On Error Resume Next
        ' Array: refNum, item, quantity, price, uom, vendor, location, itemCode, rowNum, date
        If colRefNum > 0 Then newRow.Range(1, colRefNum).value = itemData(0)    ' ref_NUMBER
        If colItems > 0 Then newRow.Range(1, colItems).value = itemData(1)      ' ITEMS
        If colQty > 0 Then newRow.Range(1, colQty).value = itemData(2)          ' QUANTITY
        If colPrice > 0 Then newRow.Range(1, colPrice).value = itemData(3)      ' PRICE
        If colUOM > 0 Then newRow.Range(1, colUOM).value = itemData(4)          ' UOM
        If colVendor > 0 Then newRow.Range(1, colVendor).value = itemData(5)    ' VENDOR
        If colLocation > 0 Then newRow.Range(1, colLocation).value = itemData(6) ' LOCATION
        If colItemCode > 0 Then newRow.Range(1, colItemCode).value = itemData(7) ' ITEM_CODE
        If colRow > 0 Then newRow.Range(1, colRow).value = itemData(8)          ' ROW
        If colEntryDate > 0 Then newRow.Range(1, colEntryDate).value = itemData(9) ' ENTRY_DATE
        On Error GoTo ErrorHandler
    Next key
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error in LogReceivedDetailed: " & Err.Description, vbCritical
End Sub

