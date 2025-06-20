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

Option Explicit
' Handle btnSend click event
Private Sub btnSend_Click()
    Call modTS_Received.ProcessReceivedBatch
    Unload Me
End Sub


Private Sub UserForm_Initialize()
   ' The lstBox should already be populated by TallyOrders()
   ' Center the form on screen
   Me.StartUpPosition = 0 'Manual
   Me.Left = Application.Left + (Application.Width - Me.Width) / 2
   Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

'────────────────────────────────────────────────────────────
' Main entry point: does all three phases: Log, update invSys RECEIVED, clear ReceivedTally and invSysData_Received
'────────────────────────────────────────────────────────────
' This will show you in the Immediate window for every row:
'    Which row you picked up
'    The values you’re about to log
'    Whether you actually called your logging routine
Public Sub ProcessReceivedBatch()
    Dim batchRef As String: batchRef = modTS_Log.GenerateOrderNumber()
    Dim lst As MSForms.ListBox: Set lst = frmReceivedTally.lstBox
    Dim i As Long, itemsLogged As Long: itemsLogged = 0

    Debug.Print "=== Starting ProcessReceivedBatch: batchRef=" & batchRef & " ==="

    For i = 0 To lst.ListCount - 1
        Dim itemName As String: itemName = CStr(lst.List(i, 0) & "")
        If itemName = "" Or itemName = "ITEMS" Then GoTo NextRow

        Dim qty    As Double: qty    = Val(lst.List(i, 1))
        Dim price  As Double: price  = Val(lst.List(i, 2))
        Dim code   As String: code   = CStr(lst.List(i, 3) & "")
        Dim rowNum As Long:   rowNum = Val(lst.List(i, 4))

        Debug.Print "Row " & i & ": item=" & itemName & "; qty=" & qty & "; price=" & price & "; code=" & code & "; row=" & rowNum

        Dim uom As String, vendor As String, location As String, entryD As Date
        GetReceivingDetails code, rowNum, uom, vendor, location, entryD
        Debug.Print " → Looked up UOM=" & uom & ", vendor=" & vendor & ", location=" & location & ", entryDate=" & entryD

        ' Log it
        Debug.Print " → Appending to ReceivedLog..."
        AppendReceivedLogRecord batchRef, itemName, qty, price, uom, vendor, location, code, rowNum, entryD
        itemsLogged = itemsLogged + 1

        ' Update INV SYS
        UpdateReceivedQuantity rowNum, qty

NextRow:
    Next i

    Debug.Print "Processed " & itemsLogged & " items; now clearing staging."
    ClearReceivedStaging
    Debug.Print "=== Done ProcessReceivedBatch ==="
End Sub


'────────────────────────────────────────────────────────────
' Pulls the rest of the staging fields from invSysData_Receiving
'────────────────────────────────────────────────────────────
Private Sub GetReceivingDetails( _
    ByVal itemCode  As String, _
    ByVal rowNum    As Long, _
    ByRef uom       As String, _
    ByRef vendor    As String, _
    ByRef location  As String, _
    ByRef entryDate As Date)
    
    With ThisWorkbook.Sheets("ReceivedTally").ListObjects("invSysData_Receiving")
        Dim lr As ListRow
        For Each lr In .ListRows
            With lr.Range
                If .Cells(.ListColumns("ROW").Index).Value = rowNum Then
                    uom       = CStr(.Cells(.ListColumns("UOM").Index).Value)
                    vendor    = CStr(.Cells(.ListColumns("VENDOR").Index).Value)
                    location  = CStr(.Cells(.ListColumns("LOCATION").Index).Value)
                    entryDate = CDate(.Cells(.ListColumns("ENTRY_DATE").Index).Value)
                    Exit Sub
                End If
            End With
        Next lr
    End With
    
    ' fallback if not found
    uom       = ""
    vendor    = ""
    location  = ""
    entryDate = Now
End Sub


'────────────────────────────────────────────────────────────
' Appends a single row into the ReceivedLog table
'────────────────────────────────────────────────────────────
Private Sub AppendReceivedLogRecord( _
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
    
    Dim ws  As Worksheet:  Set ws  = ThisWorkbook.Sheets("ReceivedLog")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("ReceivedLog")
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    With tbl.ListColumns
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


'────────────────────────────────────────────────────────────
' Adds qty to the RECEIVED column in invSys by ListObject row index
'────────────────────────────────────────────────────────────
Private Sub UpdateReceivedQuantity(ByVal rowNum As Long, ByVal qty As Double)
    Dim wsInv As Worksheet:   Set wsInv = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Dim tblInv As ListObject: Set tblInv = wsInv.ListObjects("invSys")
    
    ' rowNum corresponds 1:1 to the ListRow index
    With tblInv.ListRows(rowNum).Range
        .Cells(tblInv.ListColumns("RECEIVED").Index).Value = _
            Val(.Cells(tblInv.ListColumns("RECEIVED").Index).Value) + qty
    End With
End Sub


'────────────────────────────────────────────────────────────
' Clears both ReceivedTally and invSysData_Receiving tables
'────────────────────────────────────────────────────────────
Private Sub ClearReceivedStaging()
    Dim sht As Worksheet: Set sht = ThisWorkbook.Sheets("ReceivedTally")
    With sht.ListObjects("ReceivedTally")
        If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
    End With
    With sht.ListObjects("invSysData_Receiving")
        If Not .DataBodyRange Is Nothing Then .DataBodyRange.Delete
    End With
End Sub
'────────────────────────────────────────────────────────────
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

 Private Function GetUOMFromDataTable(item As String, ItemCode As String, rowNum As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Dim dataTbl As ListObject
    Dim uom As String
    Dim uomCol As Long, codeCol As Long, rowCol As Long
    Dim i As Long               ' ← Declare your loop counter

    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set dataTbl = ws.ListObjects("invSysData_Receiving")
    uom = "each"

    ' Find UOM column
    For i = 1 To dataTbl.ListColumns.Count
        Select Case UCase(dataTbl.ListColumns(i).Name)
            Case "UOM":        uomCol = i
            Case "ITEM_CODE":  codeCol = i
            Case "ROW":        rowCol  = i
        End Select
    Next i

    ' Search for match
    For i = 1 To dataTbl.ListRows.Count
        Dim found As Boolean
        found = False
        If rowNum <> "" And rowCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, rowCol).Value) = rowNum Then found = True
        ElseIf ItemCode <> "" And codeCol > 0 Then
            If CStr(dataTbl.DataBodyRange(i, codeCol).Value) = ItemCode Then found = True
        End If
        If found And uomCol > 0 Then
            uom = CStr(dataTbl.DataBodyRange(i, uomCol).Value)
            Exit For
        End If
    Next i

    GetUOMFromDataTable = uom
End Function



