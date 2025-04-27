VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceivedTally 
   Caption         =   "Items Received Tally"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "frmReceivedTally.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceivedTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Handle btnSend click event
Private Sub btnSend_Click()
    On Error GoTo ErrorHandler
    ' Process the data first
    SendOrderData
    ' Then unload the form
    Unload Me
    Exit Sub
ErrorHandler:
    Debug.Print "Error in btnSend_Click: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
Private Sub UserForm_Initialize()
   ' The lstBox should already be populated by TallyOrders()
   ' Center the form on screen
   Me.StartUpPosition = 0 'Manual
   Me.Left = Application.Left + (Application.Width - Me.Width) / 2
   Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub
Sub SendOrderData()
    On Error GoTo ErrorHandler
    Dim i As Long
    Dim receivedSummary As Object
    Set receivedSummary = CreateObject("Scripting.Dictionary")
    ' Create unique reference number for this batch
    Dim batchRefNumber As String
    batchRefNumber = modTS_Log.GenerateOrderNumber()
    ' Skip header row (row 0)
    For i = 1 To Me.lstBox.ListCount - 1
        Dim item As String, quantity As Double, uom As String
        Dim ItemCode As String, rowNum As String
        item = Me.lstBox.List(i, 0)             ' Item name
        quantity = CDbl(Me.lstBox.List(i, 1))   ' Quantity
        uom = Me.lstBox.List(i, 2)              ' UOM
        ItemCode = Me.lstBox.List(i, 3)         ' ItemCode (hidden column)
        rowNum = Me.lstBox.List(i, 4)           ' ROW (hidden column)
        ' Create a unique key with ROW or ITEM_CODE
        Dim uniqueKey As String
        If rowNum <> "" Then
            uniqueKey = "ROW_" & rowNum
        ElseIf ItemCode <> "" Then
            uniqueKey = "CODE_" & ItemCode
        Else
            uniqueKey = "NAME_" & item & "|" & uom
        End If
        ' Get additional data from invSysData_Receiving
        Dim price As Double, vendor As String, location As String
        GetItemDetailsFromDataTable item, ItemCode, rowNum, price, vendor, location
        ' Store complete information in dictionary
        receivedSummary(uniqueKey) = Array(batchRefNumber, item, quantity, price, uom, vendor, location, ItemCode, rowNum, Now())
    Next i
    ' Log the received items to ReceivedLog
    modTS_Log.LogReceivedDetailed receivedSummary
    ' Update quantities in inventory system
    UpdateInventory receivedSummary, "RECEIVED"
    ' Notify user
    MsgBox "Received items have been logged and inventory updated.", vbInformation
    ' Close form after processing
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
' New function to get additional details from invSysData_Receiving
Private Sub GetItemDetailsFromDataTable(itemName As String, ItemCode As String, rowNum As String, _
                                      ByRef price As Double, ByRef vendor As String, ByRef location As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Dim dataTbl As ListObject
    Set dataTbl = ws.ListObjects("invSysData_Receiving")
    If dataTbl Is Nothing Then
        Debug.Print "Data table invSysData_Receiving not found"
        Exit Sub
    End If
    ' Find matching rows in data table
    Dim i As Long, matchFound As Boolean
    matchFound = False
    ' Check columns exist
    Dim hasPrice As Boolean, hasVendor As Boolean, hasLocation As Boolean
    Dim priceCol As Long, vendorCol As Long, locationCol As Long
    Dim rowCol As Long, itemCodeCol As Long, itemNameCol As Long
    ' Find column indexes
    For i = 1 To dataTbl.ListColumns.count
        Select Case UCase(dataTbl.ListColumns(i).Name)
            Case "PRICE"
                hasPrice = True
                priceCol = i
            Case "VENDOR"
                hasVendor = True
                vendorCol = i
            Case "LOCATION"
                hasLocation = True
                locationCol = i
            Case "ROW"
                rowCol = i
            Case "ITEM_CODE"
                itemCodeCol = i
            Case "ITEMS"
                itemNameCol = i
        End Select
    Next i
    ' Initialize default values
    price = 0
    vendor = ""
    location = ""
    ' Look for matching rows in data table
    For i = 1 To dataTbl.ListRows.count
        Dim rowMatch As Boolean
        rowMatch = False
        ' Match by ROW first (most precise)
        If rowNum <> "" And rowCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, rowCol).value) = rowNum Then
                rowMatch = True
            End If
        ' Then by ITEM_CODE
        ElseIf ItemCode <> "" And itemCodeCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, itemCodeCol).value) = ItemCode Then
                rowMatch = True
            End If
        ' Finally by item name
        ElseIf itemNameCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, itemNameCol).value) = itemName Then
                rowMatch = True
            End If
        End If
        ' If we found a match, get the details
        If rowMatch Then
            matchFound = True
            ' Get PRICE
            If hasPrice Then
                On Error Resume Next
                price = price + CDbl(dataTbl.DataBodyRange(i, priceCol).value)
                On Error GoTo 0
            End If
            ' Get VENDOR (use first one found)
            If hasVendor And vendor = "" Then
                vendor = CStr(dataTbl.DataBodyRange(i, vendorCol).value)
            End If
            ' Get LOCATION (use first one found)
            If hasLocation And location = "" Then
                location = CStr(dataTbl.DataBodyRange(i, locationCol).value)
            End If
        End If
    Next i
    If Not matchFound Then
        Debug.Print "No matching rows found in data table for " & itemName
    End If
End Sub
' Function to update inventory based on ROW or ITEM_CODE
Private Sub UpdateInventory(itemsDict As Object, ColumnName As String)
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim key As Variant
    Dim foundRow As Long
    Dim currentQty As Double, newQty As Double
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    ' Get column index for the target column (e.g., "RECEIVED", "SHIPMENTS")
    Dim targetColIndex As Integer
    targetColIndex = tbl.ListColumns(ColumnName).Index
    ws.Unprotect
    Application.EnableEvents = False
    For Each key In itemsDict.Keys
        Dim itemData As Variant
        itemData = itemsDict(key)
        ' Extract info from the array
        Dim item As String, quantity As Double
        Dim ItemCode As String, rowNum As String
        item = itemData(0)
        quantity = itemData(1)
        ItemCode = itemData(3) ' itemCode at index 3
        rowNum = itemData(4)   ' rowNum at index 4
        foundRow = 0
        ' Try to find by ROW number first (most specific)
        If rowNum <> "" Then
            On Error Resume Next
            foundRow = FindRowByValue(tbl, "ROW", rowNum)
            On Error GoTo ErrorHandler
        End If
        ' If ROW didn't work, try ITEM_CODE
        If foundRow = 0 And ItemCode <> "" Then
            On Error Resume Next
            foundRow = FindRowByValue(tbl, "ITEM_CODE", ItemCode)
            On Error GoTo ErrorHandler
        End If
        ' As last resort, try finding by item name
        If foundRow = 0 Then
            On Error Resume Next
            foundRow = FindRowByValue(tbl, "ITEM", item)
            On Error GoTo ErrorHandler
        End If
        ' If we found the row, update it
        If foundRow > 0 Then
            ' Get current quantity
            currentQty = 0
            On Error Resume Next
            currentQty = tbl.DataBodyRange(foundRow, targetColIndex).value
            If IsEmpty(currentQty) Then currentQty = 0
            On Error GoTo ErrorHandler
            ' Update with new quantity
            newQty = currentQty + quantity
            tbl.DataBodyRange(foundRow, targetColIndex).value = newQty
            ' Log this change
            LogInventoryChange "UPDATE", ItemCode, item, quantity, newQty
        Else
            ' Log that we couldn't find the item
            LogInventoryChange "ERROR", ItemCode, item, quantity, 0
        End If
    Next key
    Application.EnableEvents = True
    ws.Protect
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    ws.Protect
    MsgBox "Error updating inventory: " & Err.Description, vbCritical
End Sub
' Helper function to find a row by column value
Private Function FindRowByValue(tbl As ListObject, colName As String, value As Variant) As Long
    Dim i As Long
    Dim colIndex As Integer
    FindRowByValue = 0 ' Default return value if not found
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0
    If colIndex = 0 Then Exit Function
    For i = 1 To tbl.ListRows.count
        If tbl.DataBodyRange(i, colIndex).value = value Then
            FindRowByValue = i
            Exit Function
        End If
    Next i
End Function
' Helper function to log inventory changes
Private Sub LogInventoryChange(Action As String, ItemCode As String, itemName As String, qtyChange As Double, newQty As Double)
    ' This would call your inventory logging system
    On Error Resume Next
    ' You might want to use the modTS_Log module for this
End Sub
' Add this function to frmReceivedTally.frm:
Private Function GetUOMFromDataTable(item As String, ItemCode As String, rowNum As String) As String
    On Error Resume Next
    Dim ws As Worksheet, dataTbl As ListObject
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set dataTbl = ws.ListObjects("invSysData_Receiving")
    Dim uom As String
    uom = "each" ' Default
    ' Find UOM column
    Dim uomCol As Long, itemCol As Long, codeCol As Long, rowCol As Long
    For i = 1 To dataTbl.ListColumns.count
        Select Case UCase(dataTbl.ListColumns(i).Name)
            Case "UOM": uomCol = i
            Case "ITEMS": itemCol = i
            Case "ITEM_CODE": codeCol = i
            Case "ROW": rowCol = i
        End Select
    Next i
    ' Search for match
    For i = 1 To dataTbl.ListRows.count
        Dim found As Boolean
        found = False
        If rowNum <> "" And rowCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, rowCol).value) = rowNum Then found = True
        ElseIf ItemCode <> "" And codeCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, codeCol).value) = ItemCode Then found = True
        ElseIf item <> "" And itemCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, itemCol).value) = item Then found = True
        End If
        If found And uomCol > 0 Then
            uom = CStr(dataTbl.DataBodyRange(i, uomCol).value)
            Exit For
        End If
    Next i
    GetUOMFromDataTable = uom
End Function


