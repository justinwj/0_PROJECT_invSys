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

Private Sub PopulateReceivedForm(frm As frmReceivedTally)
    Dim ws As Worksheet
    Dim inputTbl As ListObject
    Dim dataArr As Variant
    Dim idxItems As Long, idxQty As Long, idxPrice As Long
    Dim i As Long
    Dim defaultUOM As String, uom As String
    Dim itemName As String, qty As Double, prc As Double
    Dim qtyDict As Object, priceDict As Object
    Dim key As Variant

    ' Initialize
    defaultUOM = "N/A"
    Set qtyDict = CreateObject("Scripting.Dictionary")
    Set priceDict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set inputTbl = ws.ListObjects("ReceivedTally")

    ' Validate required columns
    idxItems = ColumnIndex(inputTbl, "ITEMS")
    idxQty = ColumnIndex(inputTbl, "QUANTITY")
    idxPrice = ColumnIndex(inputTbl, "PRICE")
    If idxItems * idxQty * idxPrice = 0 Then
        Err.Raise vbObjectError + 2001, , _
            "Required column missing in 'ReceivedTally': ITEMS, QUANTITY, or PRICE"
    End If

    ' Exit if no data rows
    If inputTbl.DataBodyRange Is Nothing Then
        frm.lstBox.Clear
        Exit Sub
    End If

    dataArr = inputTbl.DataBodyRange.Value

    ' Aggregate quantities and prices by item name
    For i = LBound(dataArr, 1) To UBound(dataArr, 1)
        itemName = CStr(dataArr(i, idxItems))
        qty = Val(dataArr(i, idxQty))
        prc = Val(dataArr(i, idxPrice))
        If qtyDict.Exists(itemName) Then
            qtyDict(itemName) = qtyDict(itemName) + qty
            priceDict(itemName) = priceDict(itemName) + prc
        Else
            qtyDict.Add itemName, qty
            priceDict.Add itemName, prc
        End If
    Next i

    ' Configure listbox headers
    With frm.lstBox
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "150;70;50;70"
        .AddItem "ITEMS"
        .List(0, 1) = "QUANTITY"
        .List(0, 2) = "UOM"
        .List(0, 3) = "PRICE"
    End With

    ' Populate aggregated data rows
    For Each key In qtyDict.Keys
        ' Get UOM for the item (cast key to String)
        uom = GetUOMFromInvSys(CStr(key), "", "UOM")
        If Len(Trim(uom)) = 0 Then uom = defaultUOM

        With frm.lstBox
            .AddItem CStr(key)
            .List(.ListCount - 1, 1) = qtyDict(key)
            .List(.ListCount - 1, 2) = uom
            .List(.ListCount - 1, 3) = priceDict(key)
        End With
    Next key
End Sub

' Helper: Get column index by header name
Private Function ColumnIndex(tbl As ListObject, header As String) As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If StrComp(col.Name, header, vbTextCompare) = 0 Then
            ColumnIndex = col.Index
            Exit Function
        End If
    Next col
    ColumnIndex = 0
End Function

' Helper: Get a field from invSys master by ROW
Private Function GetInvSysValue(rowNum As String, itemCode As String, header As String) As String
    Dim invWs As Worksheet, invTbl As ListObject
    Dim findCol As Long, tgtCol As Long, cel As Range
    Set invWs = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set invTbl = invWs.ListObjects("invSys")
    findCol = invTbl.ListColumns("ROW").Index
    tgtCol = invTbl.ListColumns(header).Index
    For Each cel In invTbl.DataBodyRange.Columns(findCol).Cells
        If CStr(cel.Value) = rowNum Then
            GetInvSysValue = cel.Offset(0, tgtCol - findCol).Value
            Exit Function
        End If
    Next
    GetInvSysValue = ""
End Function


' It aggregates quantities by item and displays them in a list box.
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







