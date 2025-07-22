Attribute VB_Name = "mod_Form_Load_Controls"
' Inside mod_Form_Load_Controls
Public str_Series_Name As String
Public str_Manufacturer_Name As String
Public ws_Manufacturer_Name As Worksheet
Public str_Equipment_Type As String

Public obj_Form As Object


'*********************START Initialize frm_New_Listings**************************
Public Sub init_form_New_Listings()

    Application.ScreenUpdating = False
   ' Call mod_Form_Load_Controls.load_Manufacturer_Names
    Call mod_Form_Load_Controls.load_Manufacturer_Names(frm_New_Listings)

    DoEvents
    
   ' obj_Form.lst_Manufacturer_Names.ListIndex = -1  ' Ensure no item is auto-selected
    Call mod_Create_Dictionaries.Build_Metadata          ' ? now works
    Application.ScreenUpdating = True

End Sub
'Public Sub load_Manufacturer_Names()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim cell As Range
'
'    Set ws = ThisWorkbook.Sheets("Lists")
'    Set tbl = ws.ListObjects("tbl_Manufacturer_Names")
'
'    obj_Form.lst_Manufacturer_Names.Clear
'
'
'    For Each cell In tbl.ListColumns(1).DataBodyRange
'        If Trim(cell.value) <> "" Then
'            obj_Form.lst_Manufacturer_Names.AddItem cell.value
'        End If
'    Next cell
'End Sub
'Public Sub cleanup_Binder_References()
'    Dim i As Long
'    If IsArray(fieldBinders) Then
'        For i = LBound(fieldBinders) To UBound(fieldBinders)
'            If Not fieldBinders(i) Is Nothing Then
'                Set fieldBinders(i) = Nothing
'            End If
'        Next i
'    End If
'End Sub
Public Function get_Order_Anchor_Range(ByVal orderIdentifier As Variant, _
                                      Optional ws As Worksheet = Nothing) As Range
    Dim cell As Range
    Dim lastRow As Long, i As Long

    ' Use specified sheet or default to "Orders"
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Orders")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' ?? Pass 1: Merge-aware search (robust against layout changes)
    For Each cell In ws.Range("A1:A" & lastRow)
        If cell.MergeCells Then
            If Trim(cell.MergeArea.Cells(1, 1).value) = CStr(orderIdentifier) Then
                Set get_Order_Anchor_Range = cell.MergeArea
                Exit Function
            End If
        End If
    Next cell

    ' ?? Pass 2: Stepped scan every 11 rows (optimized for fixed block layout)
    For i = 1 To lastRow Step 11
        If ws.Cells(i, "A").value = orderIdentifier Then
            Set get_Order_Anchor_Range = ws.Cells(i, "A")
            Exit Function
        End If
    Next i

    ' ? Not found
    Set get_Order_Anchor_Range = Nothing
End Function


'*********************END Initialize frm_New_Listings**************************
'Public Sub load_Manufacturer_Names(ByRef frm As Object)
'    'RUNNING CODE 7-21-2025
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim cell As Range
'
'    Set ws = ThisWorkbook.Sheets("Lists")
'    Set tbl = ws.ListObjects("tbl_Manufacturer_Names")
'
'    frm.lst_Manufacturers.Clear
'
'    For Each cell In tbl.ListColumns(1).DataBodyRange
'        If Trim(cell.value) <> "" Then
'            frm.lst_Manufacturers.AddItem cell.value
'            Debug.Print "? Manufacturer Added: " & cell.value
'        End If
'    Next cell
'End Sub

Public Sub On_Manufacturer_User_Selection(ByRef frm As Object)
    'Dim str_Manufacturer_Name As String
    'Dim ws_Manufacturer_Name As Worksheet
    Dim lastRow As Long, r As Long
    Dim valA As String


        frm.lst_Models.Clear
        frm.lst_Series.Clear
    ' ? Ensure a manufacturer is selected
    If frm.lst_Manufacturers.ListIndex = -1 Then
        MsgBox "Please select a manufacturer first.", vbExclamation
        Exit Sub
    End If

    str_Manufacturer_Name = frm.lst_Manufacturers.List(frm.lst_Manufacturers.ListIndex)
    Debug.Print "? Selected Manufacturer: [" & str_Manufacturer_Name & "]"

    ' ? Confirm the worksheet exists
    Dim wsExists As Boolean: wsExists = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = str_Manufacturer_Name Then
            wsExists = True
            Exit For
        End If
    Next ws

    If Not wsExists Then
        MsgBox "Manufacturer sheet '" & str_Manufacturer_Name & "' not found.", vbCritical
        Exit Sub
    End If

    Set ws_Manufacturer_Name = Worksheets(str_Manufacturer_Name)
    lastRow = ws_Manufacturer_Name.Cells(ws_Manufacturer_Name.Rows.Count, 1).End(xlUp).Row

    ' ? Populate Series Name list box from column A
    With frm.lst_Series
        .Clear
        For r = 3 To lastRow
            valA = Trim(ws_Manufacturer_Name.Cells(r, 1).value)
            If valA <> "" Then .AddItem valA
        Next r
    End With

    ' ? Rebuild series dictionary
    Call mod_Create_Dictionaries.create_Series_Dictionary

    ' ? Optional: capture selected series name and equipment type
    If frm.lst_Series.ListCount > 0 Then
       
        str_Series_Name = frm.lst_Series.List(0) ' Or use ListIndex if user selects
    Debug.Print "? TypeName(dict_Series): " & TypeName(dict_Series)
Debug.Print "? str_Series_Name: [" & str_Series_Name & "]"
Debug.Print "? dict_Series.Exists(str_Series_Name): " & dict_Series.Exists(str_Series_Name)

        If dict_Series.Exists(str_Series_Name) Then
    On Error Resume Next
    mod_Form_Load_Controls.str_Equipment_Type = dict_Series(str_Series_Name)

On Error Resume Next
str_Equipment_Type = dict_Series(str_Series_Name)("Equipment Type")
If Err.Number <> 0 Then
    MsgBox "Error assigning equipment type: " & Err.Description, vbCritical
    Debug.Print "? Error: " & Err.Description
    Err.Clear
Else
    Debug.Print "? Equipment Type: " & str_Equipment_Type
End If
On Error GoTo 0


    Debug.Print "? Equipment Type: " & str_Equipment_Type
Else
    MsgBox "Series '" & str_Series_Name & "' not found in dictionary.", vbExclamation
    Debug.Print "? Series not found: " & str_Series_Name
End If


    End If
End Sub





Public Sub On_Series_Name_User_Selection(ByRef frm As Object)
    Dim modelData As Object
    Dim modelName As Variant
    Dim selectedModel As String
    
    frm.lst_Models.Clear

    str_Series_Name = frm.lst_Series.value

    Call create_Model_Dictionary
        Debug.Print "? dict_Models Is Nothing: " & (dict_Models Is Nothing)
    If Not dict_Models Is Nothing Then
        Debug.Print "? dict_Models.Count: " & dict_Models.Count
    End If

    If Not dict_Models Is Nothing Then
        frm.lst_Models.Clear
        For Each modelName In dict_Models.Keys
            frm.lst_Models.AddItem modelName
        Next modelName

        If frm.lst_Models.ListIndex >= 0 Then
    selectedModel = frm.lst_Models.value
    Debug.Print "?? Selected model: " & selectedModel
Else
    Debug.Print "?? No model selected—skipping assignment."
End If


        'selectedModel = frm.lst_Models.value

        ' ? Only call Populate_InputSheet_FromModelDictionary for frm_New_Listings
        If TypeName(frm) = "frm_New_Listings" Then
            Call Populate_InputSheet_FromModelDictionary
        Else
            Debug.Print "?? Skipped Populate_InputSheet_FromModelDictionary for form: " & TypeName(frm)
        End If
    End If
    
    
    'Call DebugModelDictionaryForForm(frm_New_Orders)
    Call DumpModelDictionaryToSheetAndConsole

    
    
    
End Sub


Sub handle_chk_Hidden_Dictionaries()

    If frm_New_Listings.chk_Hidden_Dictionaries.value = True Then
        Call mod_Helper_Functions.Unhide_Dictionary_Sheets
    Else
        Call mod_Helper_Functions.Hide_Dictionary_Sheets
    End If
End Sub



'*************************Initialize Form New Orders *************************************






Public Function get_Fabric_Abbr_From_Short_Name(ByVal shortName As String) As String
    Dim abbr As Variant
    For Each abbr In dict_Fabrics.Keys
        If dict_Fabrics(abbr).Exists("Fabric Type Short Name") Then
            If StrComp(Trim(dict_Fabrics(abbr)("Fabric Type Short Name")), Trim(shortName), vbTextCompare) = 0 Then
                get_Fabric_Abbr_From_Short_Name = abbr
                Exit Function
            End If
        End If
    Next abbr

    get_Fabric_Abbr_From_Short_Name = "UNKNOWN"
End Function


Public Sub handle_lst_Fabric_Type_Names_Change(ByRef frm As Object)
    Dim selectedShortName As String
    selectedShortName = frm.lst_Fabric_Types.value

    Debug.Print "?? handle_lst_Fabric_Type_Names_Change triggered"
    Debug.Print "?? Selected short name: [" & selectedShortName & "]"
    Debug.Print "?? Map lookup? " & fabric_Display_Map.Exists(selectedShortName)

    If fabric_Display_Map.Exists(selectedShortName) Then
        Dim abbr As String
        abbr = fabric_Display_Map(selectedShortName)
        Debug.Print "?? Fabric selected ? Abbr: [" & abbr & "]"

        If Len(Trim(abbr)) = 0 Then
            Debug.Print "?? Empty abbreviation retrieved—skipping color load"
            Exit Sub
        End If

        Call populate_Color_Names(frm, abbr)
    Else
        MsgBox "Unrecognized fabric type: " & selectedShortName, vbExclamation
        frm.lst_Fabric_Colors.Clear
        Debug.Print "? Fabric type not recognized by map ? [" & selectedShortName & "]"
    End If
End Sub




Public Sub populate_Color_Names(ByRef frm As Object, ByVal fabricAbbr As String)
    frm.lst_Fabric_Colors.Clear
    Debug.Print "?? Starting color filter for fabricAbbr: [" & fabricAbbr & "]"
    
    Dim colorKey As Variant
    Dim subDict As Scripting.Dictionary
    Dim availableArray As Variant
    Dim availableRaw As Variant

    For Each colorKey In dict_Color_Names.Keys
        Debug.Print "?? ColorKey: [" & colorKey & "]"
        Set subDict = dict_Color_Names(colorKey)

        If subDict.Exists("Color Available") And subDict.Exists("My Color Name") Then
            availableRaw = subDict("Color Available")
            Debug.Print "?? Raw 'Color Available': Type = " & TypeName(availableRaw)

            Select Case TypeName(availableRaw)
                Case "String"
                    Dim cleanedRaw As String
                    cleanedRaw = Replace(Trim(CStr(availableRaw)), " ", "")
                    Debug.Print "?? Cleaned string ? [" & cleanedRaw & "]"
                    If UCase(cleanedRaw) <> "SKIP" And Len(cleanedRaw) > 0 Then
                        availableArray = Split(cleanedRaw, ",")
                        Debug.Print "?? Split array ? [" & Join(availableArray, ", ") & "]"
                    Else
                        availableArray = Array()
                        Debug.Print "?? Skipped due to SKIP or empty string"
                    End If

                Case "Variant()", "Array", "String()"
                    availableArray = availableRaw
                    Debug.Print "?? Existing array ? [" & Join(availableArray, ", ") & "]"

                Case Else
                    availableArray = Array()
                    Debug.Print "?? Unexpected 'Color Available' format ? " & TypeName(availableRaw)
            End Select

            Debug.Print "?? Test: fabricAbbr [" & fabricAbbr & "] vs list [" & Join(availableArray, ", ") & "]"
            If IsInArray(fabricAbbr, availableArray) Then
                frm.lst_Fabric_Colors.AddItem subDict("My Color Name")
                Debug.Print "? Color added: " & subDict("My Color Name")
            Else
                Debug.Print "? Color excluded: " & subDict("My Color Name") & " ? not valid for [" & fabricAbbr & "]"
            End If
        Else
            Debug.Print "?? Missing keys: 'Color Available' or 'My Color Name' in colorKey [" & colorKey & "]"
        End If
    Next colorKey

    Debug.Print "?? Finished populating colors ? Final count: " & frm.lst_Fabric_Colors.ListCount
End Sub


Public Sub On_Model_Name_User_Selection(ByRef frm As Object)
    Dim modelName As String
    modelName = frm.lst_Models.value
    Debug.Print "?? Model selected: [" & modelName & "]"

    If dict_Models Is Nothing Then
        Debug.Print "? Model dictionary not initialized."
        Exit Sub
    End If

    If Not dict_Models.Exists(modelName) Then
        Debug.Print "? Model not found in dictionary: [" & modelName & "]"
        Exit Sub
    End If

    Dim orderIndex As Long
    orderIndex = frm.GetSelectedOrderIndex
    If orderIndex = -1 Then
        Debug.Print "? No active order selection—skipping write."
        Exit Sub
    End If

    Dim startRow As Long
    startRow = GetOrderStartRow(orderIndex)
    If startRow = -1 Then
        Debug.Print "? Invalid order index [" & orderIndex & "]—cannot resolve start row."
        Exit Sub
    End If

    Dim modelDict As Scripting.Dictionary
    Set modelDict = dict_Models(modelName)

    Debug.Print "?? Dimensions for model [" & modelName & "]: " & _
                "W:" & modelDict("Width") & " D:" & modelDict("Depth") & _
                " H:" & modelDict("Height") & " OptDepth:" & modelDict("Opt. Depth")

    With ws
        .Cells(startRow + 1, 6).value = modelName                      ' Column F: Model Name
        .Cells(startRow + 2, 2).value = modelDict("Width")            ' Column B: Width
        .Cells(startRow + 2, 3).value = modelDict("Depth")            ' Column C: Depth
        .Cells(startRow + 2, 4).value = modelDict("Height")           ' Column D: Height
        .Cells(startRow + 2, 5).value = modelDict("Opt. Depth")       ' Column E: Optional Depth
    End With

    Debug.Print "? Dimensions written for [" & modelName & "] at OrderIndex [" & orderIndex & "] ? Row " & startRow + 2
End Sub




'Public Sub On_Model_Name_User_Selection(ByRef frm As Object)
'
'    Dim modelName As String
'    modelName = frm.lst_Model_Names.value
'
'    Dim orderIndex As Integer
'    orderIndex = frm.GetSelectedOrderIndex
'    If orderIndex = -1 Then
'        Debug.Print "?? No order selected—skipping dimension write."
'        Exit Sub
'    End If
'
'    Dim startRow As Long
'    startRow = GetOrderStartRow(orderIndex)
'    If startRow = -1 Then
'        Debug.Print "?? Invalid order index—cannot locate start row."
'        Exit Sub
'    End If
'    ws.Cells(startRow + 1, 6).value = frm.lst_Model_Names.value   ' Column F
'
'    If dict_Models Is Nothing Then
'        Debug.Print "?? Model dictionary not loaded."
'        Exit Sub
'    End If
'
'    If Not dict_Models.Exists(modelName) Then
'        Debug.Print "?? Selected model not found in dictionary: " & modelName
'        Exit Sub
'    End If
'
'    ' ?? Write dimensions to Orders sheet
'   ' Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Orders")
'    Dim modelDict As Object: Set modelDict = dict_Models(modelName)
'
'    With ws
'        .Cells(startRow + 2, 2).value = modelDict("Width")          ' O-Width ? Column B
'        .Cells(startRow + 2, 3).value = modelDict("Depth")          ' O-Depth ? Column C
'        .Cells(startRow + 2, 4).value = modelDict("Height")         ' O-Height ? Column D
'        .Cells(startRow + 2, 5).value = modelDict("Opt. Depth")     ' Top-Depth (Opt) ? Column E
'    End With
'
'    Debug.Print "? Dimensions written for model '" & modelName & "' at row " & (startRow + 2)
'End Sub
Public Sub refresh_Form_From_Selected_Block(frm As Object)
    Dim blockIndex As Long
    blockIndex = get_Selected_Block_Index(frm)

    If blockIndex < 0 Or blockIndex > UBound(var_Order_Type_Block_List) Then
        Debug.Print "? Invalid block index: " & blockIndex
        Exit Sub
    End If

    ' ?? Clear stale data first
    Call clear_Form_Fields(frm)

    ' ?? Load fresh data from block
    With var_Order_Type_Block_List(blockIndex)
        'frm.txt_OrderNumber.Text = .str_Order_Number
        frm.txt_Customer_Name.Text = .str_Customer_Name
        'frm.txt_Notes.Text = .str_Notes
        frm.lst_Platforms.ListIndex = -1
       ' frm.lst_Platforms.AddItem .str_Platforms
        frm.lst_Fabric_Colors.ListIndex = -1
        'frm.lst_Fabric_Colors.AddItem .str_Fabric_Colors
        'frm.lst_Manufacturers.Clear
        'frm.lst_Manufacturers.AddItem .str_Manufacturers
        frm.lst_Series.ListIndex = -1
        'frm.lst_Series.AddItem .str_Series
        frm.lst_Models.ListIndex = -1
        'frm.lst_Models.AddItem = .str_Models

        ' Add more bindings as needed
    End With
End Sub
Public Sub clear_Form_Fields(frm As Object)
    frm.boo_Is_Bulk_Clearing = True
    With frm
        '.txt_OrderNumber.Text = ""
        .txt_Customer_Name.Text = ""
        '.txt_Notes.Text = ""
        .lst_Platforms.ListIndex = -1
        .lst_Fabric_Colors.ListIndex = -1
        .lst_Manufacturers.ListIndex = -1
        .lst_Series.ListIndex = -1
        .lst_Models.ListIndex = -1
        frm.cb_Zipper_Handle.value = False
        frm.cb_Pick_Pocket.value = False
        frm.cb_Music_Rest_Design.value = False
        frm.cb_Priority_Shipping.value = False
        '.txt_Width.Text = ""
        '.txt_Depth.Text = ""
        '.txt_Height.Text = ""
        '.txt_Order_Weight.Text = ""
        ' Clear other controls as needed
    End With
    frm.boo_Is_Bulk_Clearing = False
End Sub

