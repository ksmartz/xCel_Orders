Attribute VB_Name = "Module1"
Public Sub load_Manufacturer_Names(ByRef frm As Object)
    'RUNNING CODE 7-21-2025
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Lists")
    Set tbl = ws.ListObjects("tbl_Manufacturer_Names")

    frm.lst_Manufacturers.Clear

    For Each cell In tbl.ListColumns(1).DataBodyRange
        If Trim(cell.value) <> "" Then
            frm.lst_Manufacturers.AddItem cell.value
            Debug.Print "? Manufacturer Added: " & cell.value
        End If
    Next cell
End Sub


'Public Sub Populate_Order_Sheet_From_Order_Form(frm As Object)
'    Dim orderIndex As Integer
'    Dim startRow As Long
'    Dim wsOrders As Worksheet
'    Dim modelName As String
'    Dim modelDict As Object
'
'    orderIndex = frm.GetSelectedOrderIndex
'    If orderIndex = -1 Then
'        MsgBox "No order number selected.", vbExclamation
'        Exit Sub
'    End If
'
'    startRow = GetOrderStartRow(orderIndex)
'    If startRow = -1 Then
'        MsgBox "Invalid order index.", vbCritical
'        Exit Sub
'    End If
'
'    Set wsOrders = ThisWorkbook.Worksheets("Orders")
'
'    If wsOrders.Cells(startRow, 1).value <> orderIndex Then
'        MsgBox "Mismatch in Orders sheet: Expected " & orderIndex & " in cell A" & startRow, vbCritical
'        Exit Sub
'    End If
'
'    ' ?? Get selected model name
'    modelName = frm.lst_Model_Names.value
'
'    ' ? Write basic order data (Row: startRow, Columns B–H)
'    With wsOrders
'        .Cells(startRow, 2).value = frm.txt_Customer_Name.Text
'        .Cells(startRow, 3).value = frm.lst_Platforms_Names.value
'        .Cells(startRow, 4).value = frm.lst_Fabric_Type_Names.value
'        .Cells(startRow, 5).value = frm.lst_Fabric_Color_Names.value
'        .Cells(startRow, 6).value = frm.lst_Manufacturer_Names.value
'        .Cells(startRow, 7).value = frm.lst_Series_Name.value
'        .Cells(startRow, 8).value = modelName
'    End With
'
'    ' ? If model exists, write dimensions + print-block
'    If dict_Models.Exists(modelName) Then
'        Set modelDict = dict_Models(modelName)
'
'        ' ?? Write dimensions ? Row = startRow + 2, Columns B–E
'        With wsOrders
'            .Cells(startRow + 2, 2).value = modelDict("Width")
'            .Cells(startRow + 2, 3).value = modelDict("Depth")
'            .Cells(startRow + 2, 4).value = modelDict("Height")
'            .Cells(startRow + 2, 5).value = modelDict("Opt. Depth")
'        End With
'
'        ' ?? Write temporary print-block ? Columns I–L
'        Dim colStart As Long: colStart = 9 ' Column I
'        With wsOrders
'            .Cells(startRow, colStart).value = "Customer: " & frm.txt_Customer_Name.Text
'            .Cells(startRow + 1, colStart).value = "Model: " & modelName
'            .Cells(startRow + 2, colStart).value = "Dimensions:"
'            .Cells(startRow + 2, colStart + 1).value = modelDict("Width") & """W"
'            .Cells(startRow + 2, colStart + 2).value = modelDict("Depth") & """D"
'            .Cells(startRow + 2, colStart + 3).value = modelDict("Height") & """H"
'        End With
'    Else
'        Debug.Print "?? Model not found in dictionary: " & modelName
'    End If
'
'    MsgBox "? Order data written to row " & startRow, vbInformation
'End Sub



'Public Sub RenderOrderFormLayout(targetCell As Range, modelDict As Object, frm As Object)
'    Dim ws As Worksheet: Set ws = targetCell.Worksheet
'    Dim r As Long: r = targetCell.Row
'    Dim c As Long
'
'    ' ?? Row 1 — Name, Platform, Date
'    ws.Cells(r, 2).value = "Name:"
'    ws.Cells(r, 3).value = frm.txt_Customer_Name.Text
'    ws.Cells(r, 4).value = "Platform:"
'    ws.Cells(r, 5).value = frm.lst_Platforms_Names.value
'    ws.Cells(r, 6).value = "Date:"
'    ws.Cells(r, 7).value = Date
'
'    ' ?? Row 2 — Model Info
'    ws.Cells(r + 1, 2).value = "Model Information:"
'    ws.Cells(r + 1, 3).value = modelDict("Manufacturer Name")
'    ws.Cells(r + 1, 4).value = modelDict("Series Name")
'    ws.Cells(r + 1, 5).value = modelDict("Model Name")
'    ws.Cells(r + 1, 6).value = "Equipment Type:"
'    ws.Cells(r + 1, 7).value = modelDict("Equipment Type")
'
'    ' ?? Row 3 — Fabric Info
'    ws.Cells(r + 2, 2).value = "Fabric Information:"
'    ws.Cells(r + 2, 3).value = frm.lst_Fabric_Type_Names.value
'    ws.Cells(r + 2, 4).value = "Color:"
'    ws.Cells(r + 2, 5).value = frm.lst_Fabric_Color_Names.value
'    ws.Cells(r + 2, 6).value = "Weight:"
'    ws.Cells(r + 2, 7).value = modelDict("Weight") ' assuming oz/lbs string
'
'    ' ??? Row 4 — Selected Options
'    ws.Cells(r + 3, 2).value = "Selected Options:"
'    c = 3
'    If Not frm.lst_Option_Names Is Nothing Then
'        Dim i As Long
'        For i = 0 To frm.lst_Option_Names.ListCount - 1
'            If frm.lst_Option_Names.Selected(i) Then
'                ws.Cells(r + 3, c).value = frm.lst_Option_Names.List(i)
'                c = c + 1
'            End If
'        Next i
'    End If
'
'    ' ?? Row 5 — Dimensions
'    ws.Cells(r + 4, 2).value = "Dimensions:"
'    ws.Cells(r + 4, 3).value = modelDict("Width") & """W"
'    ws.Cells(r + 4, 4).value = modelDict("Depth") & """D"
'    ws.Cells(r + 4, 5).value = modelDict("Height") & """H"
'    ws.Cells(r + 4, 6).value = IIf(Len(modelDict("Opt. Depth")) > 0, modelDict("Opt. Depth") & """D opt.", "")
'    ws.Cells(r + 4, 7).value = IIf(Len(modelDict("Opt. Height")) > 0, modelDict("Opt. Height") & """H opt.", "")
'
'    ' ?? Row 6 — Mounting or Music Rest logic
'    If modelDict("Equipment Type") = "Guitar Amp" Then
'        ws.Cells(r + 5, 2).value = "Amp Handle Location:"
'        ws.Cells(r + 5, 3).value = modelDict("AH: Location")
'        ws.Cells(r + 5, 4).value = modelDict("TAH/SAH: Length/Height") & " x " & modelDict("TAH/SAH: Width")
'        ws.Cells(r + 5, 6).value = modelDict("TAH Offset: Rear")
'        ws.Cells(r + 5, 7).value = modelDict("Angle Type")
'    ElseIf modelDict("Equipment Type") = "Music Keyboard" Then
'        ws.Cells(r + 5, 2).value = "Music Rest Design:"
'        ws.Cells(r + 5, 3).value = IIf(modelDict("Music Rest") = "Yes", "Yes", "No")
'        If modelDict("Music Rest") = "Yes" Then
'            ws.Cells(r + 5, 4).value = modelDict("Music Rest Dimensions")
'        End If
'    End If
'
'    ' ?? Row 7 — Cutting Logic
'    ws.Cells(r + 6, 3).value = "1-Piece:"
'    ws.Cells(r + 6, 4).value = "=C5+1" ' Example padding
'    ws.Cells(r + 6, 5).value = "X"
'    ws.Cells(r + 6, 6).value = "=D5+1"
'
'    ' ?? Row 11 — Notes
'    ws.Cells(r + 10, 2).value = "Notes:"
'    ws.Range(ws.Cells(r + 10, 3), ws.Cells(r + 10, 7)).Merge
'    ws.Cells(r + 10, 3).value = modelDict("General Info")
'
'    ' ?? Clean formatting
'    ws.Range(ws.Cells(r, 2), ws.Cells(r + 10, 7)).Font.Size = 10
'    ws.Range(ws.Cells(r, 2), ws.Cells(r + 10, 7)).HorizontalAlignment = xlLeft
'End Sub



'Public Sub WritePlatformName(frm As Object, value As String, keepValue As Boolean)
'    Dim r As Long: r = GetOrderStartRow(GetSelectedOrderIndex(frm))
'    If r < 1 Then Exit Sub
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("Orders")
'
'    ws.Cells(r, 4).value = IIf(keepValue, value, "") ' Column D for Platform Name
'End Sub

'Public Sub SyncOrderToSheet(frmTarget As Object, r As Long)
'
'    Dim fieldMap As Variant
'    fieldMap = GetOrderFieldMap()
'
'    If r < 1 Then Exit Sub
'
'    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Orders")
'
'    ' ?? Sync control-based fields via field map
'    Dim fieldMap As Variant: fieldMap = GetOrderFieldMap()
'    Dim i As Long, ctrlName As String, colIndex As Long
'    Dim ctrl As Object
'
'    For i = LBound(fieldMap) To UBound(fieldMap)
'        ctrlName = fieldMap(i)(0)
'        colIndex = fieldMap(i)(1)
'
'        On Error Resume Next
'        Set ctrl = frmTarget.Controls(ctrlName)
'        On Error GoTo 0
'
'        If Not ctrl Is Nothing Then
'            Select Case TypeName(ctrl)
'                Case "TextBox", "ComboBox", "Label"
'                    ws.Cells(r, colIndex).value = ctrl.value
'                Case "ListBox"
'                    If ctrl.ListIndex >= 0 Then ws.Cells(r, colIndex).value = ctrl.value
'                Case "CheckBox"
'                    ws.Cells(r, colIndex).value = IIf(ctrl.value, "Yes", "No")
'                Case "ToggleButton"
'                    ws.Cells(r, colIndex).value = IIf(ctrl.value, "On", "Off")
'            End Select
'        End If
'    Next i
'
'    ' ?? Write calculated/literal values directly
'    With ws
'        ' Row 3: Fabric weight
'        .Cells(r + 2, 7).value = str_Fabric_Weight
'
'        ' Row 5: Dimensions
'        .Cells(r + 4, 3).value = str_Model_Width        ' C5
'        .Cells(r + 4, 4).value = str_Model_Depth        ' D5
'        .Cells(r + 4, 5).value = str_Model_Height       ' E5
'        .Cells(r + 4, 6).value = str_Optional_Depth     ' F5
'        .Cells(r + 4, 7).value = str_Optional_Height    ' G5
'
'        ' Row 6: Amp or Music Rest Info
'        If str_Equipment_Type = "Guitar Amp" Then
'            .Cells(r + 5, 3).value = str_Handle_Location
'            .Cells(r + 5, 4).value = str_Handle_Dimensions
'            .Cells(r + 5, 5).value = str_Handle_Material
'            .Cells(r + 5, 6).value = str_Handle_Offset
'            .Cells(r + 5, 7).value = str_Angle_Type
'        ElseIf str_Equipment_Type = "Music Keyboard" Then
'            .Cells(r + 5, 3).value = IIf(bln_Music_Rest_Selected, "Yes", "No")
'            If bln_Music_Rest_Selected Then
'                .Cells(r + 5, 4).value = str_Music_Rest_Dimensions
'            End If
'        End If
'
'        ' Row 7: Calculated fields
'        .Cells(r + 6, 4).value = str_Calculated_1
'        .Cells(r + 6, 5).value = "X"
'        .Cells(r + 6, 6).value = str_Calculated_2
'
'        ' Row 11: Model notes
'        .Cells(r + 10, 2).value = str_Model_Notes
'    End With
'End Sub
'Public Sub Sync_Review_Order_From_To_Order_Sheet_Literals_Only(ws As Worksheet, r As Long)
'    If r < 1 Then Exit Sub
'    With ws
'        '.Cells(r, 7).value = str_Order_Date ' Row 1, Column G = Order date
'
'        .Cells(r + 1, 7).value = str_Equipment_Type ' Row 2, Column G = Equipment Type
'
'        .Cells(r + 2, 7).value = str_Fabric_Weight ' Row 3 Column G = Fabric Weight
'
'        .Cells(r + 3, 3).value = str_Model_Width 'Row 4
'        .Cells(r + 3, 4).value = str_Model_Depth 'Row 4
'        .Cells(r + 3, 5).value = str_Model_Height 'Row 4
'        .Cells(r + 3, 6).value = str_Optional_Depth 'Row 4
'        .Cells(r + 3, 7).value = str_Optional_Height 'Row 4
'
'        .Cells(r + 4, 3).value = str_Model_Width 'Row 5
'        .Cells(r + 4, 4).value = str_Model_Depth 'Row 5
'        .Cells(r + 4, 5).value = str_Model_Height 'Row 5
'        .Cells(r + 4, 6).value = str_Optional_Depth 'Row 5
'        .Cells(r + 4, 7).value = str_Optional_Height 'Row 5
'        '.Cells(r + 4, 8).value = str_Optional_Height 'Row 5
'
'        .Cells(r + 5, 3).value = str_Model_Width 'Row 6
'        .Cells(r + 5, 4).value = str_Model_Depth 'Row 6
'        .Cells(r + 5, 5).value = str_Model_Height 'Row 6
'        .Cells(r + 5, 6).value = str_Optional_Depth 'Row 6
'        .Cells(r + 5, 7).value = str_Optional_Height 'Row 6
'
'
'        .Cells(r + 6, 2).value = "" 'Row 7
'        .Cells(r + 6, 4).value = "" 'Row 7
'        .Cells(r + 6, 6).value = "" 'Row 7
'        .Cells(r + 6, 7).value = "" 'Row 7
'
'        .Cells(r + 7, 2).value = ""  'Row 8
'        .Cells(r + 7, 4).value = ""  'Row 8
'        .Cells(r + 7, 5).value = ""  'Row 8
'        .Cells(r + 7, 7).value = ""  'Row 8
'
'        .Cells(r + 8, 2).value = ""  'Row 9
'        .Cells(r + 8, 4).value = ""  'Row 9
'        .Cells(r + 8, 5).value = ""  'Row 9
'        .Cells(r + 8, 7).value = ""  'Row 9
'        .Cells(r + 8, 3).value = ""  'Row 9
'
'        .Cells(r + 9, 2).value = ""  'Row 10
'        .Cells(r + 9, 3).value = ""  'Row 10
'        .Cells(r + 9, 4).value = ""  'Row 10
'        .Cells(r + 9, 5).value = ""  'Row 10
'        .Cells(r + 9, 7).value = ""  'Row 10
'        .Cells(r + 9, 6).value = ""  'Row 10
'
'        .Cells(r + 10, 3).value = ""  'Row 11
''        .Cells(r + 10, 4).value = ""  'Row 11
''        .Cells(r + 10, 5).value = ""  'Row 11
''        .Cells(r + 10, 7).value = ""  'Row 11
'    End With
'End Sub


'Public Sub SyncFormTextToSheet(frm As Object, ctrlName As String, ws As Worksheet, r As Long, colIndex As Long)
'    Dim ctrl As Object: Set ctrl = frm.Controls(ctrlName)
'    ws.Cells(r, colIndex).value = ctrl.value
'End Sub
'Public Sub ApplyField(ByRef ctrl As Object, ByVal value As Variant, ByVal keepValue As Boolean)
'    If ctrl Is Nothing Then Exit Sub
'    If Not IsObject(ctrl) Then Exit Sub
'
'    Select Case TypeName(ctrl)
'        Case "TextBox", "ComboBox"
'            ctrl.value = IIf(keepValue, value, "")
'        Case "ListBox"
'            If ctrl.MultiSelect = fmMultiSelectSingle Then
'                If keepValue Then
'                    Debug.Print "? ListBox ? Applying value: " & value
'                    SelectListItem ctrl, value
'                Else
'                    ctrl.ListIndex = -1
'                End If
'            End If
'        Case Else
'            Debug.Print "? Unsupported control type: " & TypeName(ctrl)
'    End Select
'End Sub

'Public Function FirstOrderWithCustomer() As Long
'    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Orders")
'    Dim i As Long, r As Long
'
'    For i = 1 To 100  ' Adjust based on how many orders you support
'        r = GetOrderStartRow(i)
'        If r > 0 Then
'            If Trim(ws.Cells(r, 3).value) <> "" Then  ' Customer name assumed in column C
'                FirstOrderWithCustomer = i
'                Exit Function
'            End If
'        End If
'    Next i
'    FirstOrderWithCustomer = -1
'End Function

'    Public Function get_Review_Sync_Map() As Variant
'    Dim syncMap As Variant
'    syncMap = Array( _
'        Array("txt_Customer_Name", "chk_Name", "lbl_Name", 0, 3), _
'        Array("lst_Platforms_Names", "chk_Platform_Name", "lbl_Platform_Name", 0, 5), _
'        Array("lst_Manufacturer_Names", "chk_Manufacturer_Name", "lbl_Manufacturer_Name", 1, 3), _
'        Array("lst_Series_Name", "chk_Series_Name", "lbl_Series_Name", 1, 4), _
'        Array("lst_Model_Names", "chk_Model_Name", "lbl_Model_Name", 1, 5), _
'        Array("lst_Fabric_Type_Names", "chk_Fabric_Type_Name", "lbl_Fabric_Type_Name", 2, 3), _
'        Array("lst_Fabric_Color_Names", "chk_Color_Name", "lbl_color_Name", 2, 5) _
'            )
'
'
'
''            Array("cb_Zipper_Handle", "chk_Color_Name", "lbl_color_Name", 9, 5), _
''        Array("cb_Pick_Pocket", "chk_Color_Name", "lbl_color_Name", 9, 5), _
''        Array("cb_Music_Rest_Design", "chk_Color_Name", "lbl_color_Name", 9, 5), _
''        Array("cb_Priority_Shipping", "chk_Color_Name", "lbl_color_Name", 9, 5) _
'
'
'
'
'
'
'
'
'    get_Review_Sync_Map = syncMap
'End Function

'
'Public Sub handle_Opt_Form_New_Orders(frm As Object)
'    If frm.IsOrderNumberSelected Then
'        Dim selectedOption As String
'        selectedOption = GetSelectedOption(frm)
'        Debug.Print "Selected Option: " & selectedOption
'
'        If dict_Order_Option_Map Is Nothing Then
'            MsgBox "Option map not initialized. Please reload the form.", vbCritical
'            Exit Sub
'        End If
'
'        If dict_Order_Option_Map.Exists(selectedOption) Then
'    Dim targetCell As Range
'    Set targetCell = dict_Order_Option_Map(selectedOption)(0)
'    Set targetCell = targetCell.MergeArea.Cells(1, 1) ' ? Normalize to top of A1:A11
'
'    Dim ws As Worksheet: Set ws = targetCell.Worksheet
'
'    Dim r As Long
'    r = targetCell.MergeArea.Row ' ? Ensures consistent anchor at top of merged block
'    Debug.Print "?? TargetCell Address      : " & targetCell.MergeArea.Row
'    ' Continue building your data dictionary here...
'
'
'            ' ?? Pull correct values from redesigned form
'            Dim data As Object: Set data = CreateObject("Scripting.Dictionary")
'
'            With ws
'                ' Row 1: Name, Platform, Date
'                'Debug.Print "Name: " & data("Name")
'                data.Add "Name", .Cells(r, 3).value              ' C1
'                Debug.Print "Name: " & data("Name")
'                data.Add "Platform", .Cells(r, 5).value          ' E1
'                Debug.Print "Platform: " & data("Platform")
'                data.Add "Date", .Cells(r, 7).value              ' G1
'                Debug.Print "Date: " & data("Date")
'                ' Row 2: Manufacturer, Series, Model, Equipment Type
'                data.Add "Manufacturer", .Cells(r + 1, 3).value  ' C2
'                Debug.Print "Manufacturer: " & data("Manufacturer")
'                data.Add "Series", .Cells(r + 1, 4).value        ' D2
'                Debug.Print "Series: " & data("Series")
'                data.Add "Model", .Cells(r + 1, 5).value         ' E2
'                Debug.Print "Model: " & data("Model")
'                data.Add "EquipmentType", .Cells(r + 1, 7).value ' G2
'                Debug.Print "EquipmentType: " & data("EquipmentType")
'
'                ' Row 3: Fabric Type, Color, Weight
'                data.Add "FabricType", .Cells(r + 2, 3).value    ' C3
'                Debug.Print "FabricType: " & data("FabricType")
'                data.Add "FabricColor", .Cells(r + 2, 5).value   ' E3
'                Debug.Print "FabricColor: " & data("FabricColor")
'                data.Add "Weight", .Cells(r + 2, 7).value        ' G3
'                Debug.Print "Weight: " & data("Weight")
'
'                ' Row 4: Options (C4–G4)
'                Dim optList As Collection: Set optList = New Collection
'                Dim c As Long
'                For c = 3 To 7
'                    If Len(.Cells(r + 3, c).value) > 0 Then
'                        optList.Add .Cells(r + 3, c).value
'                    End If
'                Next c
'                data.Add "Options", optList
'
'                ' Row 5: Dimensions
'                data.Add "Dim_Width", .Cells(r + 4, 3).value     ' C5
'                data.Add "Dim_Depth", .Cells(r + 4, 4).value     ' D5
'                data.Add "Dim_Height", .Cells(r + 4, 5).value    ' E5
'                data.Add "Opt_Depth", .Cells(r + 4, 6).value     ' F5
'                data.Add "Opt_Height", .Cells(r + 4, 7).value    ' G5
'
'                ' Row 6: Mounting or Rest Details
'                data.Add "AH_Location", .Cells(r + 5, 3).value   ' C6
'                data.Add "AH_Size", .Cells(r + 5, 4).value       ' D6
'                data.Add "AH_Offset", .Cells(r + 5, 6).value     ' F6
'                data.Add "AngleType", .Cells(r + 5, 7).value     ' G6
'
'                ' Row 11: General Info / Notes
'                data.Add "GeneralInfo", .Cells(r + 10, 3).value  ' C11
'            End With
'
'            Call ShowOrderReviewForm(data, frm, targetCell)
'        End If
'
'        frm.txt_Customer_Name.Enabled = True
'    Else
'        MsgBox "Please select an order number before proceeding.", vbExclamation
'    End If
'End Sub

'Public Sub lock_User_Fields()
'    Dim lng_B As Long, lng_I As Long
'    Dim lng_rAnchor As Long
'    Dim str_Field_Name As String
'    Dim lng_Col_Index As Long
'    Dim lng_Row_Offset As Long
'    Dim lng_Target_Row As Long
'
'    For lng_B = LBound(var_Order_Type_Block_List) To UBound(var_Order_Type_Block_List)
'        lng_rAnchor = CLng(var_Order_Type_Block_List(lng_B).str_Anchor_Row)
'
'        For lng_I = LBound(var_Order_Sheet_Field_Map) To UBound(var_Order_Sheet_Field_Map)
'            If Not IsEmpty(var_Order_Sheet_Field_Map(lng_I)) Then
'                str_Field_Name = var_Order_Sheet_Field_Map(lng_I)(0)
'                lng_Col_Index = var_Order_Sheet_Field_Map(lng_I)(1)
'                lng_Row_Offset = var_Order_Sheet_Field_Map(lng_I)(2)
'                lng_Target_Row = lng_rAnchor + lng_Row_Offset
'
'                ' ?? Skip anchor and protected header zones
'                If lng_Col_Index = 1 And lng_Row_Offset = 0 Then GoTo SkipField
'                If lng_Row_Offset <= 2 And (lng_Col_Index = 1 Or lng_Col_Index = 3 Or lng_Col_Index = 5) Then GoTo SkipField
'                If lng_Row_Offset = 7 And lng_Col_Index = 1 Then GoTo SkipField
'                If lng_Row_Offset = 10 And lng_Col_Index = 1 Then GoTo SkipField
'
'                ' ?? Lock merged blocks safely (rows 9 & 10)
'                If lng_Row_Offset = 8 Or lng_Row_Offset = 9 Then
'                    With ws_Orders
'                        With .Cells(lng_Target_Row, 1): If .MergeCells Then .MergeArea.Locked = True Else .Locked = True: End With
'                        With .Cells(lng_Target_Row, 3): If .MergeCells Then .MergeArea.Locked = True Else .Locked = True: End With
'                        With .Cells(lng_Target_Row, 4): If .MergeCells Then .MergeArea.Locked = True Else .Locked = True: End With
'                        With .Cells(lng_Target_Row, 6): If .MergeCells Then .MergeArea.Locked = True Else .Locked = True: End With
'                    End With
'                    Debug.Print "?? Locked merged row " & lng_Target_Row
'                    GoTo SkipField
'                End If
'colo
'                ' ? Lock regular field
'                With ws_Orders.Cells(lng_Target_Row, lng_Col_Index)
'                    If .MergeCells Then .MergeArea.Locked = True Else .Locked = True
'                    Debug.Print "?? Locked [" & str_Field_Name & "] ? " & .Address(False, False)
'                End With
'SkipField:
'            End If
'        Next lng_I
'    Next lng_B
'
'    MsgBox "User-entry fields locked. Sheet is now protected against edits.", vbInformation
'End Sub


    
    
   '***********************NEW
    
    
    
    
    
   ' ? Route based on customer name presence
'Dim rFirst As Long
'rFirst = FirstOrderWithCustomer()
'
'
'Debug.Print "?? First customer index: " & rFirst
'Debug.Print "?? Anchor row: " & GetOrderStartRow(rFirst)
'Debug.Print "?? Value in C1: " & Sheets("Orders").Cells(1, 3).value
'
'
'If rFirst > 0 Then
'    ' ? Resolve worksheet row for that order
'    Dim rAnchor As Long
'    rAnchor = GetOrderStartRow(rFirst)
'    Debug.Print "?? Preloading Order #" & rFirst & " at row " & rAnchor
'
'    ' Show order entry form first
'    frm.Show
'
'    ' Then launch review form preloaded to the first matching order
'    Load frm_Order_Review
'    frm_Order_Review.SetOrderRow rAnchor
'    frm_Order_Review.Show
'Else
'    ' No customer data—just launch order entry form
'    frm.Show
'End If
'
'
''******************new
'    Application.ScreenUpdating = True
'
'End Sub
 '***************************'START Target Orders Worksheet for Anchor Row for Orders  Good 07-21-2025***************
'Function GetOrderStartCell(orderNumber As Long) As String
'
'    'Find Last Non-empty Row in Column A
'    lng_Last_Row = ws_Orders.Cells(ws_Orders.Rows.Count, "A").End(xlUp).Row
'
'    For lng_i = 1 To lng_Last_Row Step 11
'        If ws_Orders.Cells(lng_i, "A").value = orderNumber Then
'            GetOrderStartCell = ws_Orders.Cells(lng_i, "A").Address
'            Exit Function
'        End If
'    Next lng_i
'
'    'Reset variable for next procedure
'    lng_i = 0
'    lng_Last_Row = 0
'
'    GetOrderStartCell = "Order Not Found"
'
'
'End Function
'***************************'End Target Orders Worksheet for Anchor Row for Orders  Good 07-21-2025***************
'Public Function GetOrderStartRow(ByVal orderIndex As Integer) As Long
'
'    Dim ws As Worksheet
'
'    Dim cell As Range
'    Set ws = ThisWorkbook.Sheets("Orders")
'
'    For Each cell In ws.Range("A1:A200") ' Expand if needed
'        If cell.MergeCells Then
'            If Trim(cell.MergeArea.Cells(1, 1).value) = CStr(orderIndex) Then
'                GetOrderStartRow = cell.MergeArea.Row
'                Exit Function
'            End If
'        End If
'    Next cell
'
'    GetOrderStartRow = -1 ' Not found
'End Function
'Private Sub lst_Fabric_Color_Names_Click()
'
'    lng_R = GetOrderStartRow(GetSelectedOrderIndex(Me))
'
'    If lng_R > 0 Then
'        ThisWorkbook.Sheets("Orders").Cells(lng_R + 1, 2).value = Me.lst_Fabric_Color.value ' Column E
'    End If
'
'    ' ? Unlock fabric type and options
'    Me.lst_Manufacturer.Enabled = True
'    Me.lst_Series.Enabled = True
'    Me.lst_Model.Enabled = True
'
'End Sub
