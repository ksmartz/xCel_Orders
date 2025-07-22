Attribute VB_Name = "mod_Temp_Test"
Public gFabricDict As Scripting.Dictionary


'Sub TestFabricDict()
'    Set Globals.dict_Fabrics = Create_FabricTypeDictionary()
'    Debug.Print "Keys: " & Globals.dict_Fabrics.Count
'End Sub


'Public Function Create_FabricTypeDictionary() As Scripting.Dictionary
'    Dim d As Scripting.Dictionary
'    Set d = New Scripting.Dictionary
'    d.Add "Test", "Passed"
'    Set Create_FabricTypeDictionary = d
'End Function

'Sub TestFabric()
'    Dim result As Scripting.Dictionary
'    Set result = Create_FabricTypeDictionary()
'    Debug.Print result("Test")
'End Sub
'
'Sub TestLocalFabricDict()
'    Dim localFabricDict As Scripting.Dictionary
'    Set localFabricDict = Create_FabricTypeDictionary()
'
'    Debug.Print "Key count: " & localFabricDict.Count
'
'    Dim k As Variant
'    For Each k In localFabricDict.Keys
'        Debug.Print "Key: " & k & ", Name: " & localFabricDict(k)("Name") & ", Cost: " & localFabricDict(k)("CostPerSqInch")
'    Next k
'End Sub

'Sub TestGlobalAssignment()
'    Dim tempDict As Scripting.Dictionary
'    Set tempDict = Create_FabricTypeDictionary()
'    Debug.Print "TempDict count: " & tempDict.Count
'
'    Set Globals.dict_Fabrics = tempDict
'    Debug.Print "Global dict count: " & Globals.dict_Fabrics.Count
'End Sub

Sub Load_All_Dictionary_Metadata()
    Call Load_FabricTypeDictionary
    ' Call other loaders here
End Sub

Public Sub DumpModelDictionaryToSheetAndConsole()
    Dim wsDebug As Worksheet
    Dim modelKey As Variant, fieldName As Variant
    Dim modelDict As Object
    Dim r As Long, c As Long
    Dim fieldHeaders As Object
    Set fieldHeaders = CreateObject("Scripting.Dictionary")

    ' ?? Create or clear the debug sheet
    On Error Resume Next
    Set wsDebug = ThisWorkbook.Sheets("Model_Debug")
    If wsDebug Is Nothing Then
        Set wsDebug = ThisWorkbook.Sheets.Add
        wsDebug.Name = "Model_Debug"
    End If
    On Error GoTo 0
    wsDebug.Cells.Clear

    If dict_Models Is Nothing Then
        Debug.Print "?? dict_Models is not initialized."
        wsDebug.Cells(1, 1).value = "dict_Models is not initialized."
        Exit Sub
    End If

    If dict_Models.Count = 0 Then
        Debug.Print "?? dict_Models is empty."
        wsDebug.Cells(1, 1).value = "dict_Models is empty."
        Exit Sub
    End If

    ' ?? Collect all unique field names across all models
    For Each modelKey In dict_Models.Keys
        Set modelDict = dict_Models(modelKey)
        For Each fieldName In modelDict.Keys
            If Not fieldHeaders.Exists(fieldName) Then
                fieldHeaders.Add fieldName, fieldHeaders.Count + 2 ' Start from Column 2
            End If
        Next fieldName
    Next modelKey

    ' ??? Write header row
    wsDebug.Cells(1, 1).value = "Model Name"
    For Each fieldName In fieldHeaders.Keys
        wsDebug.Cells(1, fieldHeaders(fieldName)).value = fieldName
    Next fieldName

    ' ?? Write data rows + dump to console
    r = 2
    For Each modelKey In dict_Models.Keys
        Set modelDict = dict_Models(modelKey)

        Debug.Print "?? Model: " & modelKey
        wsDebug.Cells(r, 1).value = modelKey

        For Each fieldName In modelDict.Keys
            wsDebug.Cells(r, fieldHeaders(fieldName)).value = modelDict(fieldName)
            Debug.Print "    " & fieldName & " ? " & modelDict(fieldName)
        Next fieldName

        Debug.Print String(40, "-")
        r = r + 1
    Next modelKey

    wsDebug.Columns.AutoFit
    MsgBox "? Model dictionary debug dump complete. See sheet 'Model_Debug'.", vbInformation
End Sub

Public Sub PrintAllControlNames(frm As Object)
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        Debug.Print TypeName(ctrl), ctrl.Name
    Next ctrl
End Sub


Sub test()
    Set ws_Orders = ThisWorkbook.Worksheets("Orders")
    ws_Orders.Cells(1, 3).value = "Sam"

End Sub
Sub run_Order_Sheet_Meta_Data()

End Sub
Public Sub debug_Field_Map_Targets()
    Dim fieldMap As Variant
    Dim fieldName As String, targetCol As Long, rowOffset As Long
    Dim i As Long

    Debug.Print "?? Field Map Audit:"

    
    For i = LBound(var_Order_Sheet_Field_Map) To UBound(var_Order_Sheet_Field_Map)
        fieldMap = var_Order_Sheet_Field_Map(i)
        fieldName = fieldMap(0)
        targetCol = fieldMap(1)
        rowOffset = fieldMap(2)

        Debug.Print "? " & fieldName & " ? Column " & targetCol & _
                    " (" & Split(Cells(1, targetCol).Address, "$")(1) & ")" & _
                    ", Row Offset: " & rowOffset
    Next i
End Sub

Sub Launch_New_Orders_Form()
    Call initialize_Orders_Worksheet
    Call load_Order_Sheet_Data           ' ?? Load first
    'frm_New_Orders.Show                  ' ?? Triggers initialize_Form_Binders via form startup
End Sub
'
'Public Sub load_Order_Sheet_Data()
'    Call initialize_Orders_Worksheet
'    var_Order_Type_Block_List = get_Order_Type_Block_List()
'    var_Order_Sheet_Field_Map = get_Order_Sheet_Field_Map() ' ? ?? Type Mismatch here
'End Sub

'Public Function get_Order_Sheet_Field_Map() As Variant
'    Dim fields() As Variant
'    Dim lng_I As Long
'    ReDim fields(0 To 29)
'
'    ' Row 0 (Offset 0) – Anchor row
'    fields(0) = Array("sheet_Only_Date", 2, 0)
'    fields(1) = Array("str_Customer_Name", 4, 0)
'    fields(2) = Array("str_Platforms", 6, 0)
'
'    ' Row 1 (Offset 1) – Manufacturer / Series / Model
'    fields(3) = Array("str_Manufacturers", 2, 1)
'    fields(4) = Array("str_Series", 4, 1)
'    fields(5) = Array("str_Models", 6, 1)
'
'    ' Row 2 (Offset 2) – Fabric info
'    fields(6) = Array("str_Fabric_Types", 2, 2)
'    fields(7) = Array("str_Fabric_Colors", 4, 2)
'    fields(8) = Array("sheet_Only_Fabric_Weight", 6, 2)
'
'    ' Row 3 – Headers only ? no fields mapped
'
'    ' Row 4 (Offset 4) – Dimensional specs
'    fields(9) = Array("sheet_Only_Width", 1, 4)
'    fields(10) = Array("sheet_Only_Depth", 2, 4)
'    fields(11) = Array("sheet_Only_Height", 3, 4)
'    fields(12) = Array("sheet_Only_Depth_Opt", 4, 4)
'    fields(13) = Array("sheet_Only_Angle_Type", 5, 4)
'    fields(14) = Array("sheet_Only_Height_Opt", 6, 4)
'
'    ' Row 5 (Offset 5) – Cut dimensions
'    fields(15) = Array("sheet_Only_Cut_Width", 1, 5)
'    fields(16) = Array("sheet_Only_Cut_Depth", 2, 5)
'    fields(17) = Array("sheet_Only_Cut_Height", 3, 5)
'    fields(18) = Array("sheet_Only_Cut_Depth_Opt", 4, 5)
'    fields(19) = Array("sheet_Only_AH_Offset", 6, 5)
'
'    ' Row 6 (Offset 6) – Calculated specs
'    fields(20) = Array("sheet_Only_One_Piece_Width", 2, 6)
'    fields(21) = Array("sheet_Only_One_Piece_Depth", 4, 6)
'    fields(22) = Array("sheet_Only_One_AH_Size", 5, 6)
'    fields(23) = Array("sheet_Only_One_AH_Cut_Size", 6, 6)
'
'    ' Row 7 (Offset 7) – Options block (B–F merged)
'    fields(24) = Array("sheet_Only_Selected_Options", 2, 7)
'
'    ' Row 8 (Offset 8) – Direction block
'    fields(25) = Array("sheet_Only_1st_Direction", 3, 8)
'    fields(26) = Array("sheet_Only_2nd_Direction", 6, 8)
'
'    ' Row 9 (Offset 9) – Additional directions
'    fields(27) = Array("sheet_Only_3rd_Direction", 3, 9)
'    fields(28) = Array("sheet_Only_4th_Direction", 6, 9)
'
'    ' Row 10 (Offset 10) – Notes (B–F merged)
'    fields(29) = Array("sheet_Only_Notes", 2, 10)
'
'
'
'    get_Order_Sheet_Field_Map = fields  ' ? Return actual array of arrays
'End Function
Sub test_map()
    Dim m As Variant
    m = get_Order_Sheet_Field_Map()

    Debug.Print "TypeName: " & TypeName(m)
    Debug.Print "IsArray: " & IsArray(m)

    If IsArray(m) Then
        Dim i As Long
        For i = LBound(m) To UBound(m)
            Debug.Print "? Field[" & i & "]: " & m(i)(0) & " | Col: " & m(i)(1) & " | Offset: " & m(i)(2)
        Next i
    Else
        Debug.Print "? Failed: m is not an array"
    End If
End Sub

