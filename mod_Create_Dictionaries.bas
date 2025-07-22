Attribute VB_Name = "mod_Create_Dictionaries"
Dim dataDict As Scripting.Dictionary
Public dict_Design_Options As Scripting.Dictionary

Public dict_Fabrics As Scripting.Dictionary
Public dict_Models As Scripting.Dictionary
Public dict_Color_Names As Scripting.Dictionary
Public dict_Shipping As Scripting.Dictionary
Public dict_Miscellaneous As Scripting.Dictionary
Public dict_Series As Scripting.Dictionary
Public dict_MarketPlace_Specifics As Scripting.Dictionary
Public dict_Platforms As Scripting.Dictionary
Public dict_Fabric_Type_Short_Names As Scripting.Dictionary
Public dict_Fabric_Type_Abbr As Scripting.Dictionary
Public fabric_Display_Map As Scripting.Dictionary
Public dict_Order_Option_Map As Scripting.Dictionary





'*************Load & Optionally Dump Dictionaries*****************************
Public Sub Build_Metadata(Optional enableDump As Boolean = True)

    ' Load all metadata dictionaries
    Call create_Design_Options_Dictionary
    'Call create_Fabric_Type_Dictionary
    Call Build_Fabric_Dictionary_Transposed
   
    Call create_Color_Dictionary
    Call create_Shipping_Dictionary
    Call create_Miscellaneous_Dictionary
    Call Build_MarketPlace_Specifics_Dictionary
    ' Optional dump to consolidated sheet
    enableDump = True
    If enableDump Then
        Call mod_Helper_Functions.Dump_All_Dictionaries
    End If

    Debug.Print "Metadata build complete. Dump triggered: " & enableDump
End Sub
Public Sub Build_Form_New_Orders_Metadata(Optional enableDump As Boolean = True)
    'RUNNING 07-21-2025
    Call Build_Fabric_Dictionary_Transposed
    ' Load all metadata dictionaries
    Call create_Platform_Dictionary
'    Call create_Fabric_Type_Short_Name_Dictionary
'    Call create_Fabric_Type_Abbr_Dictionary
    Call create_Color_Dictionary
    'Call Build_Fabric_Dictionary_Transposed


End Sub

'*************Load Dictionaries*****************************
'***********
'******************************START MODEL DICTIONARY CODE******************************
' Returns a dictionary: Model Name ? Dictionary of FieldName ? Value
Public Sub create_Model_Dictionary()
   ' Dim ws As Worksheet
    Dim colStart As Long, colEnd As Long
    Dim rowDict As Scripting.Dictionary
    Dim colIndex As Long, header As Variant
    Dim lastRow As Long, r As Long
    Dim modelName As String

    Const enableDebug As Boolean = True
    Const enableDump As Boolean = True

    Set dict_Models = New Scripting.Dictionary
   ' Set ws_Manufacturer_Name = ThisWorkbook.Sheets(str_Manufacturer_Name)

    ' Locate merged header for selected series
    colStart = 1 ' Column G
    Do While ws_Manufacturer_Name.Cells(1, colStart).MergeCells
        If Trim(ws_Manufacturer_Name.Cells(1, colStart).MergeArea.Cells(1, 1).value) = str_Series_Name Then Exit Do
        colStart = colStart + ws_Manufacturer_Name.Cells(1, colStart).MergeArea.Columns.Count
    Loop
Debug.Print "? ws_Manufacturer_Name.Name: " & ws_Manufacturer_Name.Name

    If Not ws_Manufacturer_Name.Cells(1, colStart).MergeCells Then
        MsgBox "Series '" & str_Series_Name & "' not found on sheet '" & str_Manufacturer_Name & "'.", vbExclamation
        Set dict_Models = Nothing
        Exit Sub
    End If

    colEnd = colStart + ws_Manufacturer_Name.Cells(1, colStart).MergeArea.Columns.Count - 1
    
    
 
lastRow = ws_Manufacturer_Name.Cells(ws_Manufacturer_Name.Rows.Count, 2).End(xlUp).Row

    Debug.Print "? lastRow: " & lastRow

    
    
    
    
    
    
    
    
    
    
    
   ' lastRow = ws_Manufacturer_Name.Cells(ws_Manufacturer_Name.Rows.Count, colStart).End(xlUp).Row

    For r = 3 To lastRow
        modelName = Trim(ws_Manufacturer_Name.Cells(r, colStart).value)
        If modelName <> "" Then
            Set rowDict = New Scripting.Dictionary
            For colIndex = colStart To colEnd
                header = Trim(ws_Manufacturer_Name.Cells(2, colIndex).value)
                rowDict(header) = ws_Manufacturer_Name.Cells(r, colIndex).value
            Next colIndex
            dict_Models.Add modelName, rowDict

            If enableDebug Then
                Debug.Print "Model: " & modelName
                For Each header In rowDict.Keys
                    Debug.Print "   " & header & ": " & rowDict(header)
                Next header
                Debug.Print String(40, "-")
            End If
        End If
    Next r


    Debug.Print "Loaded dict_Models with " & dict_Models.Count & " keys."
End Sub



'******************************START var_Design_Options Sheet DICTIONARY CODE******************************

Public Sub create_Design_Options_Dictionary()
    Dim ws As Worksheet
    Dim subDict As Scripting.Dictionary
    Dim lastRow As Long
    Dim r As Long
    Dim str_design_Options As String, str_Design_Abbr As String, eqType As String, eqList As Variant
    Dim i As Long

    If dict_Design_Options Is Nothing Then
        Set dict_Design_Options = New Scripting.Dictionary
    Else
        dict_Design_Options.RemoveAll
    End If

    Set ws = ThisWorkbook.Sheets("var_Design_Options")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lastRow
        str_design_Options = Trim(ws.Cells(r, 1).value)        ' Column A: Design Option
        str_Design_Abbr = Trim(ws.Cells(r, 2).value)                ' Column B: Abbreviation
        eqType = Trim(ws.Cells(r, 4).value)                    ' Column D: Equipment Type

        If str_design_Options <> "" Then
            Set subDict = New Scripting.Dictionary
            subDict("Abbr") = str_Design_Abbr
            subDict("Price") = val(Replace(ws.Cells(r, 3).value, "$", ""))  ' Column C: Price
            'Helper Code
           subDict("Equipment") = SplitAndTrim(eqType)
            dict_Design_Options.Add str_design_Options, subDict
        End If
    Next r

    Debug.Print "Loaded dict_Design_Options with " & dict_Design_Options.Count & " entries."

End Sub



Sub create_Series_Dictionary()
    Dim lng_LastRow As Long
    Set dict_Series = New Scripting.Dictionary
    lng_LastRow = ws_Manufacturer_Name.Cells(ws_Manufacturer_Name.Rows.Count, 1).End(xlUp).Row


    For r = 3 To lng_LastRow
    valA = Trim(ws_Manufacturer_Name.Cells(r, 1).value) ' Series Name
    If valA <> "" Then
        Dim seriesRowDict As Scripting.Dictionary
        Set seriesRowDict = New Scripting.Dictionary

        ' Assuming Equipment Type is in column B (adjust index if needed)
        seriesRowDict("Equipment Type") = Trim(ws_Manufacturer_Name.Cells(r, 2).value)

        ' ? Optionally store other metadata here
        ' seriesRowDict("Notes") = Trim(ws_Manufacturer_Name.Cells(r, 3).Value)
        dict_Series.Add valA, seriesRowDict
        'dict_Series(valA) = seriesRowDict
    End If
Next r

If dict_Series Is Nothing Or dict_Series.Count = 0 Then
    MsgBox "dict_Series is either uninitialized or empty.", vbExclamation
    Exit Sub

End If




End Sub





'******************************End var_Design_Options Sheet DICTIONARY CODE******************************
'******************************START var_Fabric_Types Sheet DICTIONARY CODE******************************
Public Sub create_Fabric_Type_Dictionary()
    Dim ws As Worksheet
    Dim subDict As Scripting.Dictionary
    Dim r As Long
    Dim lastRow As Long
    Dim str_Fabric_Types As String, str_Fabric_Type_Abbr As String
    Dim dbl_Cost_Per_SqInch As Double

    If dict_Fabrics Is Nothing Then
        Set dict_Fabrics = New Scripting.Dictionary
    Else
        dict_Fabrics.RemoveAll
    End If

    Set ws = ThisWorkbook.Sheets("var_Fabric_Types")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        str_Fabric_Types = Trim(ws.Cells(r, 1).value)         ' Column A: Fabric Type (key)
        str_Fabric_Type_Abbr = Trim(ws.Cells(r, 2).value)     ' Column B: Abbr
        dbl_Cost_Per_SqInch = val(ws.Cells(r, 3).value)       ' Column C: Cost

        If str_Fabric_Types <> "" Then
            Set subDict = New Scripting.Dictionary
            subDict("Abbr") = str_Fabric_Type_Abbr
            subDict("CostPerSqInch") = dbl_Cost_Per_SqInch
            dict_Fabrics.Add str_Fabric_Types, subDict
        End If
    Next r

    Debug.Print "Loaded dict_Fabrics with " & dict_Fabrics.Count & " keys."

End Sub

Public Sub Build_Fabric_Dictionary_Transposed()
    'RUNNING 07-21-2025
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    Dim fieldName As String, abbr As String, shortName As String

    Dim subDict As Scripting.Dictionary
    Dim valueRaw As Variant

    Set ws = ThisWorkbook.Sheets("var_Fabric_Types")

    ' ? Reset dictionary
    If dict_Fabrics Is Nothing Then
        Set dict_Fabrics = New Scripting.Dictionary
    Else
        dict_Fabrics.RemoveAll
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    For c = 2 To lastCol
        abbr = Trim(ws.Cells(lastRow - 1, c).value)
        shortName = Trim(ws.Cells(lastRow, c).value)

        If Len(abbr) > 0 And UCase(abbr) <> "SKIP" Then
            Set subDict = New Scripting.Dictionary

            For r = 2 To lastRow - 2
                fieldName = Trim(ws.Cells(r, 1).value)
                valueRaw = ws.Cells(r, c).value

                If Len(fieldName) > 0 And UCase(valueRaw) <> "SKIP" Then
                    Select Case fieldName
                        Case "Ounces per Square Inch", "Square inches per linear yard", "Cost per Square Inch", _
                             "Ounces Per Linear Yard", "Cost /Linear Yard After Shipping", "Cost per Linear Yard"
                        subDict(fieldName) = val(valueRaw)
                    Case Else
                        subDict(fieldName) = Trim(valueRaw)
                    End Select
                End If
            Next r

            subDict("Fabric Type Abbreviation") = abbr
            subDict("Fabric Type Short Name") = shortName

            dict_Fabrics.Add abbr, subDict
            Debug.Print "? Fabric Added: " & abbr & " ? " & shortName
        Else
            Debug.Print "?? Skipped column " & c & ": abbr=" & abbr
        End If
    Next c

    Debug.Print "? Loaded dict_Fabrics with " & dict_Fabrics.Count & " entries."
End Sub




'******************************START var_Colors Sheet DICTIONARY CODE******************************
Public Sub create_Color_Dictionary()
    'RUNNING 07-21-2025
    Dim ws As Worksheet
    Dim subDict As Scripting.Dictionary
    Dim lastRow As Long
    Dim r As Long
    Dim str_ColorAbbr As String
    Dim rawFabrics As String

    If dict_Color_Names Is Nothing Then
        Set dict_Color_Names = New Scripting.Dictionary
    Else
        dict_Color_Names.RemoveAll
    End If

    Set ws = ThisWorkbook.Sheets("var_Colors")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lastRow
        str_ColorAbbr = Trim(ws.Cells(r, 2).value)        ' Column B: Color Abbr
        rawFabrics = Trim(ws.Cells(r, 4).value)           ' Column D: Available Fabrics

        If str_ColorAbbr <> "" And UCase(rawFabrics) <> "SKIP" And UCase(rawFabrics) <> "KEEP" Then
            Set subDict = New Scripting.Dictionary

            subDict("My Color Name") = Trim(ws.Cells(r, 1).value)       ' Column A
            subDict("Color Map") = Trim(ws.Cells(r, 3).value)           ' Column C
            subDict("Color Available") = SplitAndTrim(rawFabrics)       ' Column D

            dict_Color_Names.Add str_ColorAbbr, subDict
            Debug.Print "? Added Color: " & subDict("My Color Name") & " (" & str_ColorAbbr & ")"
        Else
            Debug.Print "?? Skipped row " & r & ": Abbr=" & str_ColorAbbr & ", Fabrics=" & rawFabrics
        End If
    Next r

    Debug.Print "? Loaded dict_Color_Names with " & dict_Color_Names.Count & " entries."
End Sub






'******************************End var_Colors Sheet DICTIONARY CODE******************************
'******************************START var_Shipping Costs Sheet DICTIONARY CODE******************************
Public Sub create_Shipping_Dictionary()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim weightVal As Variant, costVal As Variant

    If dict_Shipping Is Nothing Then
        Set dict_Shipping = New Scripting.Dictionary
    Else
        dict_Shipping.RemoveAll
    End If

    Set ws = ThisWorkbook.Sheets("var_Shipping")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lastRow
        weightVal = ws.Cells(r, 1).value         ' Column A: Weight
        costVal = ws.Cells(r, 2).value           ' Column B: Cost

        If IsNumeric(weightVal) And IsNumeric(costVal) Then
            dict_Shipping(CLng(weightVal)) = CDbl(costVal)
        End If
    Next r

    Debug.Print "Loaded dict_ShippingRates with " & dict_Shipping.Count & " weight-price pairs."
 
End Sub

'******************************END var_Shipping Costs Sheet DICTIONARY CODE******************************
'******************************Start var_Miscellaneous Sheet DICTIONARY CODE******************************
Public Sub create_Miscellaneous_Dictionary()
    Dim ws As Worksheet
    Dim fieldName As String, rawValue As String
    Dim valueArray As Variant, i As Long
    Dim lastRow As Long, r As Long

    If dict_Miscellaneous Is Nothing Then
        Set dict_Miscellaneous = New Scripting.Dictionary
    Else
        dict_Miscellaneous.RemoveAll
    End If

    Set ws = ThisWorkbook.Sheets("var_Miscellaneous")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        fieldName = Trim(ws.Cells(r, 1).value)   ' Column A: Key
        rawValue = Trim(ws.Cells(r, 2).value)    ' Column B: Value

        If fieldName <> "" And rawValue <> "" Then
            If InStr(rawValue, ",") > 0 Then
                valueArray = SplitAndTrim(rawValue)
                dict_Miscellaneous(fieldName) = valueArray
            ElseIf IsNumeric(rawValue) Then
                dict_Miscellaneous(fieldName) = val(rawValue)
            Else
                dict_Miscellaneous(fieldName) = rawValue
            End If
        End If
    Next r

    Debug.Print "Loaded dict_Miscellaneous with " & dict_Miscellaneous.Count & " entries."

End Sub

Public Sub Build_MarketPlace_Specifics_Dictionary()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("mp_Specifics")
    
    Dim market As String, field As String, key As String, value As String

    Dim r As Long, lastRow As Long
    Dim baseDict As Scripting.Dictionary, fieldDict As Scripting.Dictionary

    ' Initialize dictionary
    Set dict_MarketPlace_Specifics = New Scripting.Dictionary
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        market = Trim(ws.Cells(r, 1).value)
        field = Trim(ws.Cells(r, 2).value)
        key = Trim(ws.Cells(r, 3).value)
        value = Trim(ws.Cells(r, 4).value)

        If market <> "" And field <> "" And value <> "" Then
            If Not dict_MarketPlace_Specifics.Exists(market) Then
                Set baseDict = New Scripting.Dictionary
                dict_MarketPlace_Specifics.Add market, baseDict
            Else
                Set baseDict = dict_MarketPlace_Specifics(market)
            End If

            If key = "" Then
                baseDict(field) = value
            Else
                If Not baseDict.Exists(field) Then
                    Set fieldDict = New Scripting.Dictionary
                    baseDict.Add field, fieldDict
                Else
                    Set fieldDict = baseDict(field)
                End If
                fieldDict(key) = value
            End If
        End If
    Next r

    Debug.Print "? dict_MarketPlace_Specifics built with " & dict_MarketPlace_Specifics.Count & " marketplace entries."
End Sub


'******************************END var_Miscellaneous Sheet DICTIONARY CODE******************************
'Dictionaries Specifically for Form_OrderS************
'*******************************START CReate Platforms Dictionary
Sub create_Platform_Dictionary()
    'RUNNING 07-21-2025
    Set dict_Platforms = New Scripting.Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("var_Platforms")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow ' Assuming headers in row 1
        If Not dict_Platforms.Exists(ws.Cells(i, 1).value) Then
            dict_Platforms.Add ws.Cells(i, 1).value, ws.Cells(i, 2).value
        End If
    Next i
End Sub


'Public Sub create_Fabric_Type_Short_Name_Dictionary()
'    Set dict_Fabric_Type_Short_Names = New Scripting.Dictionary
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("var_Fabric_Types") ' Update if needed
'
'    Dim headerCol As Long
'    Dim lastCol As Long
'    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
'
'    Dim shortNameRow As Long: shortNameRow = 14 ' Adjust to match your row
'
'    For headerCol = 2 To lastCol
'        If Not ws.Cells(shortNameRow, headerCol).Text = "" Then
'            Dim shortName As String
'            shortName = Trim(ws.Cells(shortNameRow, headerCol).value)
'
'            If Len(shortName) > 0 And UCase(shortName) <> "SKIP" Then
'                dict_Fabric_Type_Short_Names.Add headerCol, shortName
'            End If
'        End If
'    Next headerCol
'End Sub
'
'Sub create_Fabric_Type_Abbr_Dictionary()
'    Set dict_Fabric_Type_Abbr = New Scripting.Dictionary
'
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("var_Fabric_Types")
'
'    Dim lastCol As Long
'    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
'
'    Dim abbrRow As Long: abbrRow = 13 ' Replace with actual row holding abbreviations
'
'    Dim colIndex As Long
'    For colIndex = 2 To lastCol
'        Dim abbr As String
'        abbr = Trim(ws.Cells(abbrRow, colIndex).value)
'
'        If Len(abbr) > 0 And UCase(abbr) <> "SKIP" Then
'            dict_Fabric_Type_Abbr.Add colIndex, abbr
'        End If
'    Next colIndex
'End Sub
'Public Function GetOrderFieldMap() As Variant
'    Dim fields() As Variant
'    ReDim fields(0 To 29)
'
'    ' Row 0 (Anchor row)
'    fields(0) = Array("sheet_Only_Date", 2, 0)              ' Column B
'    fields(1) = Array("txt_Customer_Name", 4, 0)            ' Column D
'    fields(2) = Array("lst_Platform_Names", 6, 0)           ' Column F
'
'    ' Row 1 (Offset 1 ? Manufacturer / Series / Model values)
'    fields(3) = Array("lst_Manufacturer_Names", 2, 1)       ' Column B
'    fields(4) = Array("lst_Series_Name", 4, 1)              ' Column D
'    fields(5) = Array("lst_Model_Names", 6, 1)              ' Column F
'
'    ' Row 2 (Offset 2 ? Fabric info)
'    fields(6) = Array("lst_Fabric_Type_Names", 2, 2)        ' Column B
'    fields(7) = Array("lst_Fabric_Colors", 4, 2)            ' Column D
'    fields(8) = Array("sheet_Only_Fabric_Weight", 6, 2)     ' Column F
'
'    ' Row 3 = Headers only ? no mapped values
'
'    ' Row 4 (Offset 4 ? Dimensional specs)
'    fields(9) = Array("sheet_Only_Width", 1, 4)             ' Column A
'    fields(10) = Array("sheet_Only_Depth", 2, 4)            ' Column B
'    fields(11) = Array("sheet_Only_Height", 3, 4)           ' Column C
'    fields(12) = Array("sheet_Only_Depth_Opt", 4, 4)        ' Column D
'    fields(13) = Array("sheet_Only_Angle_Type", 5, 4)       ' Column E
'    fields(14) = Array("sheet_Only_Height_Opt", 6, 4)       ' Column F
'
'    ' Row 5 (Offset 5 ? Cut dimensions)
'    fields(15) = Array("sheet_Only_Cut_Width", 1, 5)
'    fields(16) = Array("sheet_Only_Cut_Depth", 2, 5)
'    fields(17) = Array("sheet_Only_Cut_Height", 3, 5)
'    fields(18) = Array("sheet_Only_Cut_Depth_Opt", 4, 5)
'    fields(19) = Array("sheet_Only_AH_Offset", 6, 5)
'
'    ' Row 6 (Offset 6 ? Calculated specs)
'    fields(20) = Array("sheet_Only_One_Piece_Width", 2, 6)
'    fields(21) = Array("sheet_Only_One_Piece_Depth", 4, 6)
'    fields(22) = Array("sheet_Only_One_AH_Size", 5, 6)
'    fields(23) = Array("sheet_Only_One_AH_Cut_Size", 6, 6)
'
'    ' Row 7 (Offset 7 ? Options summary string)
'    fields(24) = Array("sheet_Only_Selected_Options", 1, 7)
'
'    ' Row 8 (Offset 8 ? Direction blocks)
'    fields(25) = Array("sheet_Only_1st_Direction", 3, 8)
'    fields(26) = Array("sheet_Only_2nd_Direction", 6, 8)
'
'    ' Row 9 (Offset 9 ? Directions 3 & 4)
'    fields(27) = Array("sheet_Only_3rd_Direction", 3, 9)
'    fields(28) = Array("sheet_Only_4th_Direction", 6, 9)
'
'    ' Row 10 (Offset 10 ? Notes)
'    fields(29) = Array("sheet_Only_Notes", 2, 10)
'
'    GetOrderFieldMap = fields
'End Function

'***************************'START Type frm_New_Orders Field Map/Dictionary - Good 07-21-2025***************
Public Function get_Order_Form_Field_Map() As Object
    'Declare Procedure Variables
    Dim dict As Object
    
    'Set Procedure Variables
    Set dict = CreateObject("Scripting.Dictionary")

    dict("str_Customer_Name") = "txt_Customer_Name"
    dict("str_Platforms") = "lst_Platforms"
    dict("str_Manufacturers") = "lst_Manufacturers"
    dict("str_Series") = "lst_Series"
    dict("str_Models") = "lst_Models"
    dict("str_Fabric_Types") = "lst_Fabric_Types"
    dict("str_Fabric_Colors") = "lst_Fabric_Colors"

    Set get_Order_Form_Field_Map = dict
End Function
'***************************'START Type frm_New_Orders Field Map/Dictionary- Good 07-21-2025***************
