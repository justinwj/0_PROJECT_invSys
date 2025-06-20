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
Attribute VB_Description = "UserForm to select shipments, populate staging tables, and send batch"

Attribute VB_Name = "frmShipmentsTally"
Option Explicit

'====================================
' UserForm: frmShipmentsTally
' Purpose: allow selecting shipments, populate staging tables, and send batch
'====================================

Private Sub btnSend_Click()
    ' Process all shipment records and update inventory and log tables
    Call modTS_Shipments.ProcessShipmentsBatch
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Center form on screen
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2

    ' Configure list box: ITEMS, QUANTITY, UOM, ITEM_CODE(hidden), ROW(hidden)
    With Me.lstBox
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "150;50;80;0;0"
        .AddItem "ITEMS"
        .List(0, 1) = "QUANTITY"
        .List(0, 2) = "UOM"
    End With

    ' Populate list box with existing ShipmentsTally entries
    PopulateShipmentsForm Me
End Sub

' Handle Enter/Tab to commit selection
Private Sub lstBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        CommitSelectionAndClose
        KeyCode = 0
    End If
End Sub

' Handle double-click to commit
Private Sub lstBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommitSelectionAndClose
End Sub

'************************************************
' Populate selection into invSysData_Shipping table
'************************************************
Private Sub CommitSelectionAndClose()
    Static isRunning As Boolean
    If isRunning Then Exit Sub
    isRunning = True

    Dim chosenValue    As String
    Dim chosenItemCode As String
    Dim chosenRowNum   As String
    Dim chosenVendor   As String
    Dim location       As String
    Dim ws             As Worksheet
    Dim tbl            As ListObject
    Dim dataTbl        As ListObject

    ' Get selection
    If Me.lstBox.ListIndex <> -1 Then
        chosenRowNum   = CStr(Me.lstBox.List(Me.lstBox.ListIndex, 4))
        chosenItemCode = CStr(Me.lstBox.List(Me.lstBox.ListIndex, 3))
        chosenValue    = CStr(Me.lstBox.List(Me.lstBox.ListIndex, 0))
        chosenVendor   = GetVendorByItem(chosenItemCode, chosenValue)
        location       = GetLocationByItem(chosenItemCode, chosenValue)
    Else
        isRunning = False
        Exit Sub
    End If

    ' Update sheet and data table
    Set ws = gSelectedCell.Worksheet
    Set tbl = ws.ListObjects("ShipmentsTally")
    Set dataTbl = ws.ListObjects("invSysData_Shipping")

    ' Determine tally row index within the table
    Dim tallyRowNum As Long
    tallyRowNum = gSelectedCell.Row - tbl.HeaderRowRange.Row

    ' Remove existing entry for this row in staging
    DeleteExistingDataForCell dataTbl, tallyRowNum

    ' Add new staging row
    Dim dataRow As ListRow
    Set dataRow = dataTbl.ListRows.Add
    ' Fill staging fields
    FillDataTableRow dataRow, _
                     GetItemUOMByRowNum(chosenRowNum, chosenItemCode, chosenValue), _
                     chosenVendor, location, chosenItemCode, chosenRowNum
    SetTallyRowNumber dataRow, tallyRowNum

    isRunning = False
End Sub

'************************************************
' Delete any existing staging for a given cell
'************************************************
Private Sub DeleteExistingDataForCell(dataTbl As ListObject, tallyRowNum As Long)
    Dim i As Long
    For i = dataTbl.ListRows.Count To 1 Step -1
        If dataTbl.DataBodyRange(i, dataTbl.ListColumns("TALLY_ROW").Index).Value = tallyRowNum Then
            dataTbl.ListRows(i).Delete
        End If
    Next i
End Sub

'*****************************************
' Fill new staging row with item details
'*****************************************
Private Sub FillDataTableRow(dataRow As ListRow, uom As String, vendor As String, location As String, _
                              ItemCode As String, rowNum As String)
    Dim tbl As ListObject
    Set tbl = dataRow.Parent
    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        Select Case UCase(tbl.ListColumns(i).Name)
            Case "UOM"
                dataRow.Range(1, i).Value = uom
            Case "VENDOR"
                dataRow.Range(1, i).Value = vendor
            Case "LOCATION"
                dataRow.Range(1, i).Value = location
            Case "ITEM_CODE"
                dataRow.Range(1, i).Value = ItemCode
            Case "ROW"
                dataRow.Range(1, i).Value = rowNum
            Case "ENTRY_DATE"
                dataRow.Range(1, i).Value = Now()
        End Select
    Next i
End Sub

'*****************************************
' Record the table row index for later use
'*****************************************
Private Sub SetTallyRowNumber(dataRow As ListRow, tallyRowNum As Long)
    dataRow.Range(1, dataRow.Parent.ListColumns("TALLY_ROW").Index).Value = tallyRowNum
End Sub

'*****************************************
' Helper: find vendor by ITEM_CODE or ITEMS
'*****************************************
Private Function GetVendorByItem(ItemCode As String, itemName As String) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("invSys")
    Dim colIdx As Long: colIdx = tbl.ListColumns("VENDOR").Index
    Dim foundRow As Long
    If ItemCode <> "" Then foundRow = FindRowByValue(tbl, "ITEM_CODE", ItemCode)
    If foundRow = 0 And itemName <> "" Then foundRow = FindRowByValue(tbl, "ITEM", itemName)
    If foundRow > 0 Then GetVendorByItem = CStr(tbl.DataBodyRange(foundRow, colIdx).Value)
End Function

'*****************************************
' Helper: find location by ITEM_CODE or ITEMS
'*****************************************
Private Function GetLocationByItem(ItemCode As String, itemName As String) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("invSys")
    Dim colIdx As Long: colIdx = tbl.ListColumns("LOCATION").Index
    Dim foundRow As Long
    If ItemCode <> "" Then foundRow = FindRowByValue(tbl, "ITEM_CODE", ItemCode)
    If foundRow = 0 And itemName <> "" Then foundRow = FindRowByValue(tbl, "ITEM", itemName)
    If foundRow > 0 Then GetLocationByItem = CStr(tbl.DataBodyRange(foundRow, colIdx).Value)
End Function

'*****************************************
' Helper: get UOM via existing global routine
'*****************************************
Private Function GetItemUOMByRowNum(rowNum As String, ItemCode As String, itemName As String) As String
    GetItemUOMByRowNum = modGlobals.GetItemUOMByRowNum(rowNum, ItemCode, itemName)
End Function

'*****************************************
' Helper: find a row by matching a column value
'*****************************************
Private Function FindRowByValue(tbl As ListObject, colName As String, value As Variant) As Long
    Dim i As Long, colIndex As Long
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    If colIndex = 0 Then Exit Function
    For i = 1 To tbl.ListRows.Count
        If CStr(tbl.DataBodyRange(i, colIndex).Value) = CStr(value) Then
            FindRowByValue = i: Exit Function
        End If
    Next i
End Function

