VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShipmentsTally 
   Caption         =   "Shipments Tally"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "frmShipmentsTally.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShipmentsTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Handle btnSend click event
Private Sub btnSend_Click()
    SendOrderData
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    ' Center the form on screen
    Me.StartUpPosition = 0 'Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub
Sub SendOrderData()
    On Error GoTo ErrorHandler
    Dim i As Long, shipmentsSummary As Object
    Set shipmentsSummary = CreateObject("Scripting.Dictionary")
    ' Skip header row (row 0)
    For i = 1 To Me.lstBox.ListCount - 1
        Dim item As String, quantity As Double, uom As String
        Dim ItemCode As String, rowNum As String
        item = Me.lstBox.List(i, 0)
        quantity = CDbl(Me.lstBox.List(i, 1))
        uom = Me.lstBox.List(i, 2)
        ' Get the hidden columns with ITEM_CODE and ROW
        ItemCode = Me.lstBox.List(i, 3)
        rowNum = Me.lstBox.List(i, 4)
        ' Create a unique key with ROW or ITEM_CODE
        Dim uniqueKey As String
        If rowNum <> "" Then
            uniqueKey = "ROW_" & rowNum
        ElseIf ItemCode <> "" Then
            uniqueKey = "CODE_" & ItemCode
        Else
            uniqueKey = "NAME_" & item & "|" & uom
        End If
        ' Store in dictionary with all needed information
        shipmentsSummary(uniqueKey) = Array(item, quantity, uom, ItemCode, rowNum)
    Next i
    ' Log the shipment in the ShipmentsLog
    modTS_Log.LogShipments shipmentsSummary
    ' Update quantities in inventory system
    UpdateInventory shipmentsSummary, "SHIPMENTS"
    ' Close form after processing
    Unload Me
    Exit Sub
ErrorHandler:
   MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
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

