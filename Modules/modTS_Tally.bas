Attribute VB_Name = "modTS_Tally"
' ================================================
' Module: modTS_Tally (TS stands for Tally System)
' ================================================
Option Explicit
' This module is responsible for tallying orders and displaying them in a user form.
' Track if we're already running a tally operation
Private isRunningTally As Boolean
' Helper function to normalize text
Private Function NormalizeText(text As String) As String
    ' Trim and convert to lowercase for consistent matching
    Dim result As String
    result = Trim(text)
    NormalizeText = LCase(result)
End Function
Sub TallyShipments()
    ' Create and show form with shipments data
    Dim frm As frmShipmentsTally
    Set frm = New frmShipmentsTally
    ' Make sure the form has required controls
    If Not FormHasRequiredControls(frm) Then
        MsgBox "The form is missing required controls.", vbCritical
        Exit Sub
    End If
    ' Configure the form
    With frm
        ' Make sure the listbox exists and is configured properly
        .lstBox.Clear
        .lstBox.ColumnCount = 3
        .lstBox.ColumnWidths = "150;50;80" ' Adjust as needed
        .lstBox.AddItem "ITEMS"
        .lstBox.List(0, 1) = "QUANTITY"
        .lstBox.List(0, 2) = "UOM"
    End With
    ' Populate the form
    PopulateShipmentsForm frm
    ' Show the form if there are items
    If frm.lstBox.ListCount > 1 Then ' More than just the header row
        frm.Show vbModal
    Else
        MsgBox "No shipments to tally", vbInformation
    End If
End Sub
Function FormHasRequiredControls(frm As Object) As Boolean
    On Error Resume Next
    FormHasRequiredControls = Not (frm.lstBox Is Nothing)
    On Error GoTo 0
End Function
Sub PopulateShipmentsForm(frm As frmShipmentsTally)
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    Dim itemInfo As Variant
    ' Get worksheet and table references
    Set ws = ThisWorkbook.Sheets("ShipmentsTally")
    Set tbl = ws.ListObjects("ShipmentsTally")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    ' Process and tally items from the table
    For i = 1 To tbl.ListRows.count
        ' Get basic values with error handling
        Dim item As String, quantity As Double, uom As String
        On Error Resume Next
        item = CStr(tbl.DataBodyRange(i, tbl.ListColumns("ITEMS").Index).value)
        ' Be extra careful with quantity conversion
        Dim rawQuantity As Variant
        rawQuantity = tbl.DataBodyRange(i, tbl.ListColumns("QUANTITY").Index).value
        If IsNumeric(rawQuantity) Then
            quantity = CDbl(rawQuantity)
        Else
            quantity = 0
        End If
        uom = CStr(tbl.DataBodyRange(i, tbl.ListColumns("UOM").Index).value)
        On Error GoTo ErrorHandler
        ' Skip empty rows or rows with zero quantity
        If Trim(item) <> "" And quantity > 0 Then
            ' Extract ROW and ITEM_CODE if available
            Dim rowNum As String, ItemCode As String
            rowNum = ""
            ItemCode = ""
            On Error Resume Next
            ' Check if ROW and ITEM_CODE are in columns
            For j = 1 To tbl.ListColumns.count
                If UCase(tbl.ListColumns(j).Name) = "ROW" Then
                    rowNum = CStr(tbl.DataBodyRange(i, j).value)
                ElseIf UCase(tbl.ListColumns(j).Name) = "ITEM_CODE" Then
                    ItemCode = CStr(tbl.DataBodyRange(i, j).value)
                End If
            Next j
            ' If we don't have a ROW yet, look up the item in inventory
            If rowNum = "" Then
                Dim invWs As Worksheet
                Dim invTbl As ListObject
                Dim lookupRow As Long
                Set invWs = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
                Set invTbl = invWs.ListObjects("invSys")
                If ItemCode <> "" Then
                    lookupRow = FindRowByValue(invTbl, "ITEM_CODE", ItemCode)
                End If
                If lookupRow = 0 Then
                    lookupRow = FindRowByValue(invTbl, "ITEM", item)
                End If
                If lookupRow > 0 Then
                    rowNum = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("ROW").Index).value)
                End If
            End If
            ' Create a unique key - FIXED: For shipments from inventory, ensure items from different rows stay separate
            Dim uniqueKey As String
            If rowNum <> "" Then
                ' Use ROW for uniqueness (most specific)
                uniqueKey = "ROW_" & rowNum
            ElseIf ItemCode <> "" Then
                ' Use ITEM_CODE as fallback
                uniqueKey = "CODE_" & ItemCode
            Else
                ' If no ROW or ITEM_CODE, treat each entry as unique by including row position
                uniqueKey = "NAME_" & LCase(Trim(item)) & "|" & LCase(Trim(uom)) & "|" & i
            End If
            ' Tally items using the unique key
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
            Else
                dict.Add uniqueKey, quantity
                ' Store reference information
                dict.Add "info_" & uniqueKey, Array(item, ItemCode, rowNum, uom)
            End If
        End If
    Next i
    ' Configure form list box
    frm.lstBox.Clear
    frm.lstBox.ColumnCount = 5 ' ITEM, QTY, UOM, ITEM_CODE, ROW
    frm.lstBox.ColumnWidths = "150;50;50;0;0" ' Hide ITEM_CODE and ROW
    ' Add header row
    frm.lstBox.AddItem "ITEMS"
    frm.lstBox.List(0, 1) = "QTY"
    frm.lstBox.List(0, 2) = "UOM"
    ' Add data rows
    If dict.count > 0 Then
        For Each key In dict.Keys
            If Left$(key, 5) <> "info_" Then
                itemInfo = dict("info_" & key)
                frm.lstBox.AddItem itemInfo(0) ' Item name
                frm.lstBox.List(frm.lstBox.ListCount - 1, 1) = dict(key)   ' Quantity
                frm.lstBox.List(frm.lstBox.ListCount - 1, 2) = itemInfo(3) ' UOM
                frm.lstBox.List(frm.lstBox.ListCount - 1, 3) = itemInfo(1) ' ITEM_CODE
                frm.lstBox.List(frm.lstBox.ListCount - 1, 4) = itemInfo(2) ' ROW
            End If
        Next key
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error in PopulateShipmentsForm: " & Err.Description, vbCritical
    Debug.Print "Error in PopulateShipmentsForm: " & Err.Description
    Resume Next
End Sub
Sub PopulateReceivedForm(frm As frmReceivedTally)
    On Error GoTo ErrorHandler
    ' Debug info - show what we're doing
    Debug.Print "PopulateReceivedForm: Starting to populate form..."
    Dim ws As Worksheet
    Dim tbl As ListObject, dataTbl As ListObject
    Dim dict As Object
    Dim priceDict As Object
    Dim uomDict As Object
    Dim i As Long, j As Long, k As Long
    Dim key As Variant
    Dim itemInfo As Variant
    ' Get worksheet and table references
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set tbl = ws.ListObjects("ReceivedTally")
    Set dataTbl = ws.ListObjects("invSysData_Receiving")
    ' Check if table exists
    If tbl Is Nothing Then
        MsgBox "Error: ReceivedTally table not found!", vbCritical
        Exit Sub
    End If
    ' DEBUG: Print column names to check actual structure
    Debug.Print "Table columns in ReceivedTally:"
    For i = 1 To tbl.ListColumns.count
        Debug.Print i & ": " & tbl.ListColumns(i).Name
    Next i
    ' Verify required columns exist
    Dim itemsColIndex As Long, qtyColIndex As Long, uomColIndex As Long, priceColIndex As Long
    itemsColIndex = 0: qtyColIndex = 0: uomColIndex = 0: priceColIndex = 0
    For i = 1 To tbl.ListColumns.count
        Select Case UCase(tbl.ListColumns(i).Name)
            Case "ITEMS": itemsColIndex = i
            Case "QUANTITY": qtyColIndex = i
            Case "UOM": uomColIndex = i
            Case "PRICE": priceColIndex = i
        End Select
    Next i
    ' Exit if required columns are missing
    If itemsColIndex = 0 Then
        MsgBox "Required column 'ITEMS' not found in ReceivedTally table", vbExclamation
        Exit Sub
    End If
    If qtyColIndex = 0 Then
        MsgBox "Required column 'QUANTITY' not found in ReceivedTally table", vbExclamation
        Exit Sub
    End If
    ' Create dictionaries
    Set dict = CreateObject("Scripting.Dictionary")     ' For quantities
    Set priceDict = CreateObject("Scripting.Dictionary") ' For prices
    Set uomDict = CreateObject("Scripting.Dictionary")   ' For UOMs
    dict.CompareMode = vbTextCompare
    ' Process items from invSysData_Receiving to get UOMs and prices
    Dim dataItemsColIndex As Long, dataUOMColIndex As Long, dataPriceColIndex As Long
    Dim dataItemCodeColIndex As Long, dataRowColIndex As Long
    dataItemsColIndex = 0: dataUOMColIndex = 0: dataPriceColIndex = 0
    dataItemCodeColIndex = 0: dataRowColIndex = 0
    ' Get column indexes in data table
    If Not dataTbl Is Nothing Then
        For i = 1 To dataTbl.ListColumns.count
            Select Case UCase(dataTbl.ListColumns(i).Name)
                Case "ITEMS": dataItemsColIndex = i
                Case "UOM": dataUOMColIndex = i
                Case "PRICE": dataPriceColIndex = i
                Case "ITEM_CODE": dataItemCodeColIndex = i
                Case "ROW": dataRowColIndex = i
            End Select
        Next i
        ' Build lookup dictionaries from data table
        Dim dataUOMByItem As Object, dataPriceByItem As Object
        Set dataUOMByItem = CreateObject("Scripting.Dictionary")
        Set dataPriceByItem = CreateObject("Scripting.Dictionary")
        For i = 1 To dataTbl.ListRows.count
            Dim dataItem As String, dataUOM As String, dataPrice As Double
            Dim dataItemCode As String, dataRow As String
            On Error Resume Next
            If dataItemsColIndex > 0 Then dataItem = CStr(dataTbl.DataBodyRange(i, dataItemsColIndex).value)
            If dataUOMColIndex > 0 Then dataUOM = CStr(dataTbl.DataBodyRange(i, dataUOMColIndex).value)
            If dataPriceColIndex > 0 Then
                If IsNumeric(dataTbl.DataBodyRange(i, dataPriceColIndex).value) Then
                    dataPrice = CDbl(dataTbl.DataBodyRange(i, dataPriceColIndex).value)
                End If
            End If
            If dataItemCodeColIndex > 0 Then dataItemCode = CStr(dataTbl.DataBodyRange(i, dataItemCodeColIndex).value)
            If dataRowColIndex > 0 Then dataRow = CStr(dataTbl.DataBodyRange(i, dataRowColIndex).value)
            On Error GoTo ErrorHandler
            ' Store by ROW first (most precise), ITEM_CODE second, item name last
            If dataRow <> "" Then
                If Not dataUOMByItem.Exists("ROW_" & dataRow) Then dataUOMByItem.Add "ROW_" & dataRow, dataUOM
                If Not dataPriceByItem.Exists("ROW_" & dataRow) Then dataPriceByItem.Add "ROW_" & dataRow, dataPrice
            End If
            If dataItemCode <> "" Then
                If Not dataUOMByItem.Exists("CODE_" & dataItemCode) Then dataUOMByItem.Add "CODE_" & dataItemCode, dataUOM
                If Not dataPriceByItem.Exists("CODE_" & dataItemCode) Then dataPriceByItem.Add "CODE_" & dataItemCode, dataPrice
            End If
            If dataItem <> "" Then
                If Not dataUOMByItem.Exists("NAME_" & LCase(Trim(dataItem))) Then dataUOMByItem.Add "NAME_" & LCase(Trim(dataItem)), dataUOM
                If Not dataPriceByItem.Exists("NAME_" & LCase(Trim(dataItem))) Then dataPriceByItem.Add "NAME_" & LCase(Trim(dataItem)), dataPrice
            End If
        Next i
    End If
    ' Process and tally items from the table
    For i = 1 To tbl.ListRows.count
        On Error Resume Next
        ' Use verified column indexes
        Dim item As String, quantity As Double, uom As String, price As Double
        item = CStr(tbl.DataBodyRange(i, itemsColIndex).value)
        Dim rawQuantity As Variant
        rawQuantity = tbl.DataBodyRange(i, qtyColIndex).value
        If IsNumeric(rawQuantity) Then
            quantity = CDbl(rawQuantity)
        Else
            quantity = 0
        End If
        ' Get UOM from table first
        If uomColIndex > 0 Then
            uom = CStr(tbl.DataBodyRange(i, uomColIndex).value)
        End If
        ' Get price if available
        If priceColIndex > 0 Then
            If IsNumeric(tbl.DataBodyRange(i, priceColIndex).value) Then
                price = CDbl(tbl.DataBodyRange(i, priceColIndex).value)
            End If
        End If
        ' Skip empty or zero quantity items
        If Trim(item) <> "" And quantity > 0 Then
            ' Get ROW and ITEM_CODE from inventory
            Dim rowNum As String, ItemCode As String
            rowNum = "": ItemCode = ""
            ' Get items directly from inventory
            Dim invWs As Worksheet
            Dim invTbl As ListObject
            Dim lookupRow As Long
            Set invWs = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
            Set invTbl = invWs.ListObjects("invSys")
            lookupRow = FindRowByValue(invTbl, "ITEM", item)
            If lookupRow > 0 Then
                On Error Resume Next
                rowNum = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("ROW").Index).value)
                ItemCode = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("ITEM_CODE").Index).value)
                ' Get UOM from inventory if not set yet
                If Trim(uom) = "" Then
                    uom = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("UOM").Index).value)
                End If
                On Error GoTo ErrorHandler
            End If
            ' Create a unique key for tallying
            Dim uniqueKey As String
            If rowNum <> "" Then
                uniqueKey = "ROW_" & rowNum
            ElseIf ItemCode <> "" Then
                uniqueKey = "CODE_" & ItemCode
            Else
                uniqueKey = "NAME_" & LCase(Trim(item))
            End If
            ' Tally items
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
                priceDict(uniqueKey) = priceDict(uniqueKey) + price  ' Just add prices directly
            Else
                dict.Add uniqueKey, quantity
                priceDict.Add uniqueKey, price  ' Store price without multiplication
                ' Get UOM from data table if available
                If dataUOMByItem.Exists(uniqueKey) And Trim(uom) = "" Then
                    uom = dataUOMByItem(uniqueKey)
                ElseIf Trim(uom) = "" Then
                    uom = "each" ' Default only if no other UOM found
                End If
                ' Store row information and UOM
                dict.Add "info_" & uniqueKey, Array(item, ItemCode, rowNum, uom)
                ' Store UOM for this key
                uomDict.Add uniqueKey, uom
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    ' Configure form list box
    frm.lstBox.Clear
    frm.lstBox.ColumnCount = 6 ' ITEM, QTY, UOM, PRICE, ITEM_CODE, ROW
    frm.lstBox.ColumnWidths = "150;50;50;70;0;0" ' Hide ITEM_CODE and ROW
    ' Add header row
    frm.lstBox.AddItem "ITEMS"
    frm.lstBox.List(0, 1) = "QTY"
    frm.lstBox.List(0, 2) = "UOM"
    frm.lstBox.List(0, 3) = "PRICE"
    ' Add data rows
    If dict.count > 0 Then
        For Each key In dict.Keys
            If Left$(key, 5) <> "info_" Then
                itemInfo = dict("info_" & key)
                Dim unitPrice As Double
                ' Calculate unit price (price per unit)
                If dict(key) > 0 Then
                    ' First try to get price from priceDict
                    If priceDict.Exists(key) Then
                        unitPrice = priceDict(key) / dict(key)
                    ' Then from data table lookup
                    ElseIf dataPriceByItem.Exists(key) Then
                        unitPrice = dataPriceByItem(key)
                    Else
                        unitPrice = 0
                    End If
                End If
                frm.lstBox.AddItem itemInfo(0) ' Item name
                frm.lstBox.List(frm.lstBox.ListCount - 1, 1) = dict(key)      ' Quantity
                frm.lstBox.List(frm.lstBox.ListCount - 1, 2) = itemInfo(3)   ' UOM
                frm.lstBox.List(frm.lstBox.ListCount - 1, 3) = priceDict(key)  ' Display total price
                frm.lstBox.List(frm.lstBox.ListCount - 1, 4) = itemInfo(1)   ' ITEM_CODE
                frm.lstBox.List(frm.lstBox.ListCount - 1, 5) = itemInfo(2)   ' ROW
            End If
        Next key
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error in PopulateReceivedForm: " & Err.Description & " (Line: " & Erl & ")", vbCritical
    Debug.Print "Error in PopulateReceivedForm: " & Err.Description & " at line " & Erl
    Resume Next
End Sub
Sub TallyReceived()
    On Error GoTo ErrorHandler
    ' Debug info
    Debug.Print "TallyReceived: Starting..."
    ' Verify the worksheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ReceivedTally")
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        MsgBox "The worksheet 'ReceivedTally' does not exist!", vbExclamation
        Exit Sub
    End If
    ' Verify the table exists
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects("ReceivedTally")
    On Error GoTo ErrorHandler
    If tbl Is Nothing Then
        MsgBox "The table 'ReceivedTally' does not exist on worksheet 'ReceivedTally'!", vbExclamation
        Exit Sub
    End If
    ' Create and show form with received items data
    Dim frm As New frmReceivedTally
    ' Configure the form
    With frm
        ' Make sure the listbox exists and is configured properly
        .lstBox.Clear
        .lstBox.ColumnCount = 5  ' ITEM, QTY, UOM, ITEM_CODE(hidden), ROW(hidden)
        .lstBox.ColumnWidths = "150;50;50;0;0" ' Hide ITEM_CODE and ROW columns
        .lstBox.AddItem "ITEMS"
        .lstBox.List(0, 1) = "QUANTITY"
        .lstBox.List(0, 2) = "UOM"
    End With
    ' Populate form directly using our PopulateReceivedForm function
    PopulateReceivedForm frm
    ' Show the form if there are items
    If frm.lstBox.ListCount > 1 Then ' More than just the header row
        frm.Show vbModal
    Else
        MsgBox "No received items to tally", vbInformation
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error in TallyReceived: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Debug.Print "Error in TallyReceived: " & Err.Description & " (Error " & Err.Number & ")"
End Sub
' This should be in your ribbon callback or worksheet button
Public Sub LaunchShipmentsTally()
    Application.ScreenUpdating = False
    TallyShipments
    Application.ScreenUpdating = True
End Sub
' This should be in your ribbon callback or worksheet button
Public Sub LaunchReceivedTally()
    Application.ScreenUpdating = False
    TallyReceived
    Application.ScreenUpdating = True
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
        ' Convert both values to strings for more reliable comparison
        If CStr(tbl.DataBodyRange(i, colIndex).value) = CStr(value) Then
            FindRowByValue = i
            Debug.Print "Found match in " & colName & " column: " & value & " at row " & i
            Exit Function
        End If
    Next i
    Debug.Print "No match found in " & colName & " column for value: " & CStr(value)
End Function







