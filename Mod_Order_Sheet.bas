Attribute VB_Name = "Mod_Order_Sheet"
'Public order_Number_Starting_Cells As Object
Public var_Order_Type_Block_List() As order_Information_Block
Public var_Order_Sheet_Field_Map As Variant
Public IsFormShuttingDown As Boolean

Public ws_Orders As Worksheet
Dim lng_i As Long
'Dim var_m As Variant
Dim lng_B As Long
Dim lng_RAnchor As Long
Dim lng_Last_Row As Long
'***************************'START Type Block for Order Sheet - Good 07-21-2025***************
Public Type order_Information_Block
    'RUNNING CODE 7-21-2025
    str_Order_Number As String
    str_Anchor_Row As String
    str_Customer_Name As String
    str_Platforms As String
    str_Manufacturers As String
    str_Series As String
    str_Models As String
    str_Fabric_Types As String
    str_Fabric_Colors As String
    str_Order_Weight As Double
    str_Width As Double
    str_Depth As Double
    str_Height As Double
    str_Notes As String
End Type
'***************************'End Type Block for Order Sheet - Good 07-21-2025***************
'***************************'START Set Modular Variables - Good 07-21-2025***************
Sub initialize_Orders_Worksheet()
    Set ws_Orders = ThisWorkbook.Worksheets("Orders")
End Sub
'***************************'End Set Modular Variables - Good 07-21-2025***************
'***************************'START Create Array for Order_Sheet Types - Good 07-21-2025***************
Public Function get_Order_Type_Block_List() As order_Information_Block()

    'Declare Procedural Variables
    Dim rng_Cell As Range
    Dim lng_rIndex As Long
    Dim anchorsList() As order_Information_Block
    Dim infoBlock As order_Information_Block
    Dim lng_RAnchor As Long
    
    'Set procedure Variables
    'lng_RAnchor = rng_Cell.Row

    ReDim anchorsList(0 To 0)

    For Each rng_Cell In ws_Orders.Range("A1:A1000")
        If IsNumeric(rng_Cell.value) And rng_Cell.value > 0 Then
            'Set procedure Variables
            lng_RAnchor = rng_Cell.Row
            If lng_rIndex > 0 Then ReDim Preserve anchorsList(0 To lng_rIndex)
            
            infoBlock.str_Anchor_Row = CStr(lng_RAnchor)
            infoBlock.str_Order_Number = CStr(rng_Cell.value)
    
            ' Row Offset 0
            infoBlock.str_Customer_Name = CStr(ws_Orders.Cells(lng_RAnchor + 0, 4).value)  ' Column D
            infoBlock.str_Platforms = CStr(ws_Orders.Cells(lng_RAnchor + 0, 6).value)       ' Column F
    
            ' Row Offset 1
            infoBlock.str_Manufacturers = CStr(ws_Orders.Cells(lng_RAnchor + 1, 2).value)
            infoBlock.str_Series = CStr(ws_Orders.Cells(lng_RAnchor + 1, 4).value)
            infoBlock.str_Models = CStr(ws_Orders.Cells(lng_RAnchor + 1, 6).value)
    
            ' Row Offset 2
            infoBlock.str_Fabric_Types = CStr(ws_Orders.Cells(lng_RAnchor + 2, 2).value)
            infoBlock.str_Fabric_Colors = CStr(ws_Orders.Cells(lng_RAnchor + 2, 4).value)
            infoBlock.str_Order_Weight = val(ws_Orders.Cells(lng_RAnchor + 2, 6).value)
    
            ' Row Offset 4
            infoBlock.str_Width = val(ws_Orders.Cells(lng_RAnchor + 4, 1).value)
            infoBlock.str_Depth = val(ws_Orders.Cells(lng_RAnchor + 4, 2).value)
            infoBlock.str_Height = val(ws_Orders.Cells(lng_RAnchor + 4, 3).value)
    
            ' Row Offset 10 — Notes (merged block B–F)
            infoBlock.str_Notes = val(ws_Orders.Cells(lng_RAnchor + 10, 2).value)
    
            anchorsList(lng_rIndex) = infoBlock
            lng_rIndex = lng_rIndex + 1
        End If
    Next rng_Cell

    For lng_i = LBound(anchorsList) To UBound(anchorsList)
        With anchorsList(lng_i)
            Debug.Print "?? Block[" & lng_i & "] ? AnchorRow: " & .str_Anchor_Row & _
                        ", Order #: " & .str_Order_Number & _
                        ", Customer: " & .str_Customer_Name
        End With
    Next lng_i
    
    'Reset variable for next procedure
    lng_i = 0

    get_Order_Type_Block_List = anchorsList

End Function
'***************************'End Create Array for Order_Sheet Types - Good 07-21-2025***************

'***************************'START Scan frm_New_Orders -Find opt_Button for Order # Good 07-21-2025***************
Public Function get_Selected_Block_Index(frm As Object) As Long
    'Declare Procedural Variables
    Dim controlName As String
    Dim foundIndex As Long: foundIndex = -1

    'Debug.Print "?? Scanning for selected opt_OrderX button on form: " & frm.Name

    For lng_i = 1 To 10  ' Update this range if you have more/fewer buttons
        controlName = "opt_Order" & lng_i

        If frm.Controls(controlName).value = True Then
            foundIndex = lng_i - 1
            'Debug.Print "? Button selected: " & controlName & " ? Index " & foundIndex
            Exit For
        Else
            'Debug.Print "… Not selected: " & controlName
        End If
    Next lng_i

    If foundIndex = -1 Then
        'Debug.Print "?? No opt_OrderX button is selected."
    End If
    
    'Reset variable for next procedure
    lng_i = 0

    get_Selected_Block_Index = foundIndex
End Function
'***************************'End Scan frm_New_Orders -Find opt_Button for Order # Good 07-21-2025***************

'***************************'START Clear Order Sheet # Good 07-21-2025***************
Sub clear_Order_Sheet_Routine()

    'Call sets Order Sheet
    Call initialize_Orders_Worksheet
    
    'Call sets Order Sheet
    Call load_Order_Sheet_Data  ' ?? Load blockList + fieldMap first
    
    Call clear_Complete_Order_Sheet
    'Call prompt_Clear_Order_Sheet 'This will call the sub to clear the order sheet
End Sub
'***************************'END Clear Order Sheet # Good 07-21-2025***************

'***************************'START Clear Order Sheet Routing # Good 07-21-2025***************
Public Sub load_Order_Sheet_Data()
    
    'Pull structure Order Blocks from Order Sheet
    'Populates var_Order_Type_Block_List with an array of order_Information_Block entries.
    var_Order_Type_Block_List = get_Order_Type_Block_List()
    
    'Get field map array and inspects it
    var_m = get_Order_Sheet_Field_Map()

    Debug.Print "TypeName: " & TypeName(var_m)
    Debug.Print "IsArray: " & IsArray(var_m)

    If IsArray(var_m) Then
        For lng_i = LBound(var_m) To UBound(var_m)
            Debug.Print "? Field[" & lng_i & "]: " & var_m(lng_i)(0) & " | Col: " & var_m(lng_i)(1) & " | Offset: " & var_m(lng_i)(2)
        Next lng_i
    Else
        Debug.Print "? Failed: var_m is not an array"
    End If
    
    'Reset variable for next procedure
    lng_i = 0
    
    var_Order_Sheet_Field_Map = get_Order_Sheet_Field_Map()
End Sub


'***************************'START Clear Order Sheet Routing # Good 07-21-2025***************
Public Sub prompt_Clear_Order_Sheet()
    Dim userResponse As VbMsgBoxResult

    userResponse = MsgBox("Do you want to clear the data on the Order Sheet?", vbQuestion + vbYesNo, "Confirm Data Clear")

    If userResponse = vbYes Then
        Call clear_Complete_Order_Sheet  ' ?? Step 1: clear mapped input fields
        Call lock_User_Fields            ' ?? Step 2: lock all cleared fields
        MsgBox "Order sheet cleared and locked. User-entry zones are now protected.", vbInformation, "Completed"
    Else
        MsgBox "Order sheet NOT cleared. No changes made.", vbInformation, "Cancelled"
        ' Optional: Call show_Order_Sheet_Status
    End If
End Sub

'***************************'START Clear Order Sheet Completely - Good 07-21-2025***************
Public Sub clear_Complete_Order_Sheet()
    For lng_B = LBound(var_Order_Type_Block_List) To UBound(var_Order_Type_Block_List)
        lng_RAnchor = CLng(var_Order_Type_Block_List(lng_B).str_Anchor_Row)
        
        For lng_i = LBound(var_Order_Sheet_Field_Map) To UBound(var_Order_Sheet_Field_Map)
            If Not IsEmpty(var_Order_Sheet_Field_Map(lng_i)) Then
                    Dim str_Field_Name As String
                    Dim lng_Col_Index As Long
                    Dim lng_Row_Offset As Long
                    Dim lng_Target_Row As Long
                    
                    str_Field_Name = var_Order_Sheet_Field_Map(lng_i)(0)
                    lng_Col_Index = var_Order_Sheet_Field_Map(lng_i)(1)
                    lng_Row_Offset = var_Order_Sheet_Field_Map(lng_i)(2)

                    lng_Target_Row = lng_RAnchor + lng_Row_Offset

                    ' ? Row 9 & 10 override — clear merged blocks
                    If lng_Row_Offset = 8 Or lng_Row_Offset = 9 Then
                        With ws_Orders
                            ' A–B merged
                            With .Cells(lng_Target_Row, 1)
                                If .MergeCells Then .MergeArea.ClearContents Else .ClearContents
                            End With
                            ' C
                            With .Cells(lng_Target_Row, 3)
                                If .MergeCells Then .MergeArea.ClearContents Else .ClearContents
                            End With
                            ' D–E merged
                            With .Cells(lng_Target_Row, 4)
                                If .MergeCells Then .MergeArea.ClearContents Else .ClearContents
                            End With
                            ' F
                            With .Cells(lng_Target_Row, 6)
                                If .MergeCells Then .MergeArea.ClearContents Else .ClearContents
                            End With
                        End With

                        Debug.Print "?? Cleared merged-safe row " & lng_Target_Row & " (Offset " & lng_Row_Offset & ")"
                        GoTo SkipField
                    End If

                    ' ??? Skip anchor cell (Column A at offset 0)
                    If lng_Col_Index = 1 And lng_Row_Offset = 0 Then GoTo SkipField

                    ' ??? Skip headers in rows 0–2 (Columns A, C, E)
                    If lng_Row_Offset <= 2 And (lng_Col_Index = 1 Or lng_Col_Index = 3 Or lng_Col_Index = 5) Then GoTo SkipField

                    ' ??? Skip "Options:" header at Row 7, Column A
                    If lng_Row_Offset = 7 And lng_Col_Index = 1 Then GoTo SkipField

                    ' ??? Skip "Notes:" header at Row 10, Column A
                    If lng_Row_Offset = 10 And lng_Col_Index = 1 Then GoTo SkipField

                    ' ? Clear regular mapped field
                    With ws_Orders.Cells(lng_Target_Row, lng_Col_Index)
                        If .MergeCells Then .MergeArea.ClearContents Else .ClearContents
                    Debug.Print "?? Cleared [" & str_Field_Name & "] ? " & _
                                ws_Orders.Cells(lng_Target_Row, lng_Col_Index).Address(False, False)
                    End With
SkipField:
                End If
            Next lng_i

            Debug.Print "? Finished clearing block at anchor row " & lng_RAnchor
    Next lng_B
    

    MsgBox "Order sheet cleared: all mapped fields wiped, merged and header-safe.", vbInformation
End Sub

'***************************'START Initialize frm_New_Orders Form # Good 07-21-2025***************
Public Sub init_Form_New_Orders(ByRef frm As Object)

    'Declare Procedure Variables
    Dim var_abbr As Variant
    Dim var_key As Variant
    Dim col_Platform_List As Collection
    Dim var_Sorted_Platforms As Variant
    
    'Set procedure Variable
    Set col_Platform_List = New Collection
    

    Application.ScreenUpdating = False
    
   ' Call initialize_Orders_Worksheet 'Set ws_Orders worksheet
    Call initialize_Orders_Worksheet
    Call load_Order_Sheet_Data
    'Load all meta data needed to use for the Order Sheets.
   ' Call clear_Order_Sheet_Routine
    'Call load_Order_Sheet_Data  ' ?? Load blockList + fieldMap first
   ' Call prompt_Clear_Order_Sheet 'This will call the sub to clear the order sheet
    'frm_New_Orders.Show
    
  '  Set dict_Order_Option_Map = lock_All_Form_Controls(frm)
  clear_Complete_Order_Sheet
'Call frm_New_Orders.initialize_Form_Binders
    Call load_Manufacturer_Names(frm)

    ' ? Load all metadata dictionaries
    Call mod_Create_Dictionaries.Build_Form_New_Orders_Metadata

    ' ? Clear list boxes before populating
    frm.lst_Fabric_Types.Clear
    frm.lst_Fabric_Colors.Clear
    frm.lst_Platforms.Clear

    ' ? Initialize and populate fabric_Display_Map
    Set fabric_Display_Map = New Scripting.Dictionary

    Debug.Print "?? dict_Fabrics count: " & dict_Fabrics.Count

    
    For Each abbr In dict_Fabrics.Keys
        Debug.Print "? Checking abbr: [" & abbr & "]"
        
        Debug.Print "?? Evaluating shortName: [" & shortName & "]"
Debug.Print "? Length: " & Len(shortName) & ", Upper: " & UCase(shortName)


        If dict_Fabrics(abbr).Exists("Fabric Type Short Name") Then
     
            shortName = Trim(dict_Fabrics(abbr)("Fabric Type Short Name"))

            If Len(shortName) > 0 And UCase(shortName) <> "SKIP" Then
                frm.lst_Fabric_Types.AddItem shortName
                fabric_Display_Map(shortName) = abbr
                Debug.Print "? Fabric Type Added: " & shortName & " ? " & abbr
            Else
                Debug.Print "?? Skipped short name: " & shortName
            End If
        Else
            Debug.Print "? Missing 'Fabric Type Short Name' for abbr: " & abbr
        End If
    Next abbr

    ' ? Load and alphabetize all colors
    Call load_All_Colors(frm)

   


    For Each var_key In dict_Platforms.Keys
        col_Platform_List.Add var_key
    Next var_key

    
    var_Sorted_Platforms = SortVariantArray(CollectionToArray(col_Platform_List))


    For lng_i = LBound(var_Sorted_Platforms) To UBound(var_Sorted_Platforms)
        frm.lst_Platforms.AddItem var_Sorted_Platforms(lng_i)
        Debug.Print "? Platform Added: " & var_Sorted_Platforms(lng_i)
    Next lng_i
   ' frm.Show


    Application.ScreenUpdating = True
    

    End Sub
'***************************'END Initialize frm_New_Orders Form # Good 07-21-2025***************

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


'***************************'START Set Field Map as Array  Good 07-21-2025**********************
Public Function get_Order_Sheet_Field_Map() As Variant

    Dim fields() As Variant

    ReDim fields(0 To 29)

    ' Row 0 (Offset 0) – Anchor row
    fields(0) = Array("sheet_Only_Date", 2, 0)
    fields(1) = Array("str_Customer_Name", 4, 0)
    fields(2) = Array("str_Platforms", 6, 0)

    ' Row 1 (Offset 1) – Manufacturer / Series / Model
    fields(3) = Array("str_Manufacturers", 2, 1)
    fields(4) = Array("str_Series", 4, 1)
    fields(5) = Array("str_Models", 6, 1)

    ' Row 2 (Offset 2) – Fabric info
    fields(6) = Array("str_Fabric_Types", 2, 2)
    fields(7) = Array("str_Fabric_Colors", 4, 2)
    fields(8) = Array("sheet_Only_Fabric_Weight", 6, 2)

    ' Row 3 – Headers only ? no fields mapped

    ' Row 4 (Offset 4) – Dimensional specs
    fields(9) = Array("sheet_Only_Width", 1, 4)
    fields(10) = Array("sheet_Only_Depth", 2, 4)
    fields(11) = Array("sheet_Only_Height", 3, 4)
    fields(12) = Array("sheet_Only_Depth_Opt", 4, 4)
    fields(13) = Array("sheet_Only_Angle_Type", 5, 4)
    fields(14) = Array("sheet_Only_Height_Opt", 6, 4)

    ' Row 5 (Offset 5) – Cut dimensions
    fields(15) = Array("sheet_Only_Cut_Width", 1, 5)
    fields(16) = Array("sheet_Only_Cut_Depth", 2, 5)
    fields(17) = Array("sheet_Only_Cut_Height", 3, 5)
    fields(18) = Array("sheet_Only_Cut_Depth_Opt", 4, 5)
    fields(19) = Array("sheet_Only_AH_Offset", 6, 5)

    ' Row 6 (Offset 6) – Calculated specs
    fields(20) = Array("sheet_Only_One_Piece_Width", 2, 6)
    fields(21) = Array("sheet_Only_One_Piece_Depth", 4, 6)
    fields(22) = Array("sheet_Only_One_AH_Size", 5, 6)
    fields(23) = Array("sheet_Only_One_AH_Cut_Size", 6, 6)

    ' Row 7 (Offset 7) – Options block (B–F merged)
    fields(24) = Array("sheet_Only_Selected_Options", 2, 7)

    ' Row 8 (Offset 8) – Direction block
    fields(25) = Array("sheet_Only_1st_Direction", 3, 8)
    fields(26) = Array("sheet_Only_2nd_Direction", 6, 8)

    ' Row 9 (Offset 9) – Additional directions
    fields(27) = Array("sheet_Only_3rd_Direction", 3, 9)
    fields(28) = Array("sheet_Only_4th_Direction", 6, 9)

    ' Row 10 (Offset 10) – Notes (B–F merged)
    fields(29) = Array("sheet_Only_Notes", 2, 10)

    get_Order_Sheet_Field_Map = fields  ' ? Return actual array of arrays
    
End Function
'***************************'End Set Field Map as Array  Good 07-21-2025**********************








Public Sub handle_Selected_Options_Order_Block(frm As Object)
    'RUNNING CODE 7-21-2025
    Dim blockIndex As Long
    blockIndex = get_Selected_Block_Index(frm)

    If blockIndex < 0 Or blockIndex > UBound(var_Order_Type_Block_List) Then
        MsgBox "No valid order block selected.", vbExclamation
        Exit Sub
    End If

    Dim rAnchor As Long: rAnchor = CLng(var_Order_Type_Block_List(blockIndex).str_Anchor_Row)
     ' Assumes target sheet is Orders

    Dim data As Object: Set data = CreateObject("Scripting.Dictionary")
    Dim fieldMap As Variant
    Dim fieldName As String, colIndex As Long, rowOffset As Long

    For Each fieldMap In var_Order_Sheet_Field_Map
        fieldName = fieldMap(0)
        colIndex = fieldMap(1)
        rowOffset = fieldMap(2)

        With ws_Orders.Cells(rAnchor + rowOffset, colIndex)
            data(fieldName) = IIf(.MergeCells, .MergeArea.Cells(1, 1).value, .value)
        End With
    Next fieldMap

    ' ?? Optional: log key fields
    Debug.Print "? Review Data for Block Index " & blockIndex
    Debug.Print "Name: " & data("str_Customer_Name")
    Debug.Print "Platform: " & data("str_Platforms")
    Debug.Print "Anchor Row: " & rAnchor

    ' ?? Launch review form
    'Load frm_Order_Review
    'frm_Order_Review.SetOrderRow rAnchor
   ' frm_Order_Review.Show
End Sub


Public Sub unlock_User_Fields()
    Dim lng_B As Long, lng_i As Long
    Dim lng_RAnchor As Long
    Dim str_Field_Name As String
    Dim lng_Col_Index As Long
    Dim lng_Row_Offset As Long
    Dim lng_Target_Row As Long

    For lng_B = LBound(var_Order_Type_Block_List) To UBound(var_Order_Type_Block_List)
        lng_RAnchor = CLng(var_Order_Type_Block_List(lng_B).str_Anchor_Row)

        For lng_i = LBound(var_Order_Sheet_Field_Map) To UBound(var_Order_Sheet_Field_Map)
            If Not IsEmpty(var_Order_Sheet_Field_Map(lng_i)) Then
                str_Field_Name = var_Order_Sheet_Field_Map(lng_i)(0)
                lng_Col_Index = var_Order_Sheet_Field_Map(lng_i)(1)
                lng_Row_Offset = var_Order_Sheet_Field_Map(lng_i)(2)
                lng_Target_Row = lng_RAnchor + lng_Row_Offset

                ' ?? Skip anchor cell and headers
                If lng_Col_Index = 1 And lng_Row_Offset = 0 Then GoTo SkipField
                If lng_Row_Offset <= 2 And (lng_Col_Index = 1 Or lng_Col_Index = 3 Or lng_Col_Index = 5) Then GoTo SkipField
                If lng_Row_Offset = 7 And lng_Col_Index = 1 Then GoTo SkipField
                If lng_Row_Offset = 10 And lng_Col_Index = 1 Then GoTo SkipField

                ' ?? Unlock merged blocks safely (rows 9 & 10)
                If lng_Row_Offset = 8 Or lng_Row_Offset = 9 Then
                    With ws_Orders
                        With .Cells(lng_Target_Row, 1): If .MergeCells Then .MergeArea.Locked = False Else .Locked = False: End With
                        With .Cells(lng_Target_Row, 3): If .MergeCells Then .MergeArea.Locked = False Else .Locked = False: End With
                        With .Cells(lng_Target_Row, 4): If .MergeCells Then .MergeArea.Locked = False Else .Locked = False: End With
                        With .Cells(lng_Target_Row, 6): If .MergeCells Then .MergeArea.Locked = False Else .Locked = False: End With
                    End With
                    Debug.Print "?? Unlocked merged row " & lng_Target_Row
                    GoTo SkipField
                End If

                ' ? Unlock standard field
                With ws_Orders.Cells(lng_Target_Row, lng_Col_Index)
                    If .MergeCells Then .MergeArea.Locked = False Else .Locked = False
                    Debug.Print "?? Unlocked [" & str_Field_Name & "] ? " & .Address(False, False)
                End With
SkipField:
            End If
        Next lng_i
    Next lng_B

    MsgBox "User-entry fields unlocked. Sheet ready for input.", vbInformation
End Sub

Private Function GetOptionMapFromForm(frm As Object) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    Dim ctrl As Control
    Dim index As Long

    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Parent.Name = "fra_Order_Numbers" Then
                index = CLng(Replace(ctrl.Name, "opt_Order", "")) ' Extract numeric index
                With Sheets("Orders")
    dict.Add ctrl.Name, Array(.Cells(10, index).Address, .Range(.Cells(11, index), .Cells(15, index)).Address)
End With

            End If
        End If
    Next ctrl

    Set GetOptionMapFromForm = dict


End Function
Public Sub ClearAllMappedFieldsUsingOrderMap()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Orders")
    Dim cell As Range
    Dim rAnchor As Long
    Dim fieldMap As Variant
    Dim i As Long

    fieldMap = GetOrderFieldMap()

    For Each cell In ws.Range("A1:A200")  ' Expand range as needed
        If cell.MergeCells Then
            rAnchor = cell.MergeArea.Row

            For i = LBound(fieldMap) To UBound(fieldMap)
                If Not IsEmpty(fieldMap(i)) Then
                    Dim col As Long
                    col = fieldMap(i)(1)  ' Column index from map
                    ws.Cells(rAnchor, col).ClearContents
                End If
            Next i

            Debug.Print "?? Cleared mapped fields at row " & rAnchor
        End If
    Next cell

    MsgBox "Mapped fields cleared for all order blocks.", vbInformation
End Sub


Public Sub ClearField(ctrl As Object)
    Select Case TypeName(ctrl)
        Case "TextBox", "Label", "ComboBox"
            ctrl.value = ""
        Case "ListBox"
            ctrl.Clear
        Case "CheckBox", "ToggleButton"
            ctrl.value = False
    End Select
End Sub
Public Function GetSelectedOrderIndex(frm As Object) As Long
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.value = True And ctrl.Name Like "opt_Order#" Then
                GetSelectedOrderIndex = CLng(Replace(ctrl.Name, "opt_Order", ""))
                Exit Function
            End If
        End If
    Next ctrl
    GetSelectedOrderIndex = -1 ' Not found
End Function

Public Sub populate_OrderBlock_From_Order_Form(ByRef block As order_Information_Block)
    Dim map As Object
    Set map = get_Order_Form_Field_Map()
    Dim field As Variant, controlName As String
    Dim ctl As MSForms.Control

    For Each field In map.Keys
        controlName = map(field)
        Set ctl = frm_New_Orders.Controls(controlName)

        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Or TypeName(ctl) = "ListBox" Then
            If field Like "str_*Weight" Or field Like "str_*Width" Or field Like "str_*Height" Or field = "str_Notes" Then
                CallByName block, field, VbLet, val(ctl.Text)
            Else
                CallByName block, field, VbLet, ctl.Text
            End If
        End If
    Next field
End Sub

Public Sub write_Order_Block_From_Form()
    Dim blockIndex As Long, i As Long
    Dim targetBlock As order_Information_Block
    Dim lng_RAnchor As Long
    Dim str_Field_Name As String
    Dim lng_Col_Index As Long
    Dim lng_Row_Offset As Long
    Dim lng_Target_Row As Long

    ' ?? Determine selected opt_OrderX index
    For i = 1 To 10
        If frm_New_Orders.Controls("opt_Order" & i).value = True Then
            blockIndex = i - 1
            Exit For
        End If
    Next i

    If blockIndex < 0 Or blockIndex > UBound(var_Order_Type_Block_List) Then
        MsgBox "No order slot selected. Please choose one.", vbExclamation
        Exit Sub
    End If

    ' ?? Build filled order block from form input
    targetBlock = get_OrderBlock_From_Form()
    targetBlock.str_Order_Number = frm_New_Orders.txt_OrderNumber.Text
    targetBlock.str_Anchor_Row = var_Order_Type_Block_List(blockIndex).str_Anchor_Row
    var_Order_Type_Block_List(blockIndex) = targetBlock

    lng_RAnchor = CLng(targetBlock.str_Anchor_Row)

    ' ?? Write each mapped field to its cell
    For i = LBound(var_Order_Sheet_Field_Map) To UBound(var_Order_Sheet_Field_Map)
        If Not IsEmpty(var_Order_Sheet_Field_Map(i)) Then
            str_Field_Name = var_Order_Sheet_Field_Map(i)(0)
            lng_Col_Index = var_Order_Sheet_Field_Map(i)(1)
            lng_Row_Offset = var_Order_Sheet_Field_Map(i)(2)
            lng_Target_Row = lng_RAnchor + lng_Row_Offset

            Dim cellValue As Variant
            cellValue = CallByName(targetBlock, str_Field_Name, VbGet)

            With ws_Orders.Cells(lng_Target_Row, lng_Col_Index)
                If .MergeCells Then .MergeArea.value = cellValue Else .value = cellValue
            End With

            Debug.Print "?? Wrote [" & str_Field_Name & "] ? " & _
                        ws_Orders.Cells(lng_Target_Row, lng_Col_Index).Address(False, False)
        End If
    Next i

    MsgBox "Order block written to sheet at anchor row " & lng_RAnchor, vbInformation
End Sub
Public Sub liveWrite_FieldValue(strFieldName As String, newValue As Variant)
    Dim i As Long, blockIndex As Long, lng_RAnchor As Long
    Dim lng_Col_Index As Long, lng_Row_Offset As Long, lng_Target_Row As Long

    ' ?? Find active block index (opt_OrderX)
    For i = 1 To 10
        If frm_New_Orders.Controls("opt_Order" & i).value = True Then
            blockIndex = i - 1
            Exit For
        End If
    Next i

    If blockIndex < 0 Or blockIndex > UBound(var_Order_Type_Block_List) Then Exit Sub

    lng_RAnchor = CLng(var_Order_Type_Block_List(blockIndex).str_Anchor_Row)

    ' ?? Find field map for target field
    For i = LBound(var_Order_Sheet_Field_Map) To UBound(var_Order_Sheet_Field_Map)
        If var_Order_Sheet_Field_Map(i)(0) = strFieldName Then
            lng_Col_Index = var_Order_Sheet_Field_Map(i)(1)
            lng_Row_Offset = var_Order_Sheet_Field_Map(i)(2)
            lng_Target_Row = lng_RAnchor + lng_Row_Offset

            With ws_Orders.Cells(lng_Target_Row, lng_Col_Index)
                If .MergeCells Then .MergeArea.value = newValue Else .value = newValue
            End With

            Exit Sub
        End If
    Next i
End Sub

Public Sub filter_Fabric_Colors_By_FabricType(frm As frm_New_Orders)
    Dim selectedTypeLabel As String
    Dim selectedAbbr As String
    Dim colorKey As Variant
    Dim subDict As Scripting.Dictionary
    Dim availableTypes As Variant
    Dim displayColor As String

    ' ?? Read selected fabric type
    selectedTypeLabel = frm.lst_Fabric_Types.value
    selectedAbbr = ExtractAbbreviationFromLabel(selectedTypeLabel)

    frm.lst_Fabric_Colors.Clear

    For Each colorKey In dict_Color_Names.Keys
        Set subDict = dict_Color_Names(colorKey)

        If subDict.Exists("Color Available") And subDict.Exists("My Color Name") Then
            availableTypes = subDict("Color Available")  ' This should be an array of abbrs

            If IsAbbreviationPresent(selectedAbbr, availableTypes) Then
                displayColor = subDict("My Color Name") & " (" & colorKey & ")"
                frm.lst_Fabric_Colors.AddItem displayColor
                Debug.Print "? Included Color: " & displayColor
            Else
                Debug.Print "? Hidden due to unsupported abbr: " & subDict("My Color Name")
            End If
        End If
    Next colorKey
End Sub
Private Function ExtractAbbreviationFromLabel(label As String) As String
    ' Assumes format: "Choice Fabric (C)"
    Dim startPos As Long, endPos As Long

    startPos = InStrRev(label, "(") + 1
    endPos = InStrRev(label, ")")

    If startPos > 0 And endPos > startPos Then
        ExtractAbbreviationFromLabel = Mid(label, startPos, endPos - startPos)
    Else
        ExtractAbbreviationFromLabel = ""
    End If
End Function
Private Function IsAbbreviationPresent(abbr As String, abbrList As Variant) As Boolean
    Dim i As Long
    For i = LBound(abbrList) To UBound(abbrList)
        If Trim(UCase(abbrList(i))) = Trim(UCase(abbr)) Then
            IsAbbreviationPresent = True
            Exit Function
        End If
    Next i
    IsAbbreviationPresent = False
End Function

Public Sub handle_Btn_New_Orders()
    Unload frm_Dash_Board
   ' frm_New_Orders.Show vbModeless
   ' Call frm_New_Orders.initialize_Form_Binders
    Call init_Form_New_Orders(frm_New_Orders)
    frm_New_Orders.Show
End Sub
