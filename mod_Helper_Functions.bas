Attribute VB_Name = "mod_Helper_Functions"
Public Function SplitAndTrim(ByVal rawText As String) As Variant
    Dim parts As Variant, i As Long
    If Trim(rawText) = "" Then
        SplitAndTrim = Array()
        Exit Function
    End If
    parts = Split(rawText, ",")
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim(parts(i))
    Next i
    SplitAndTrim = parts
End Function

'*********************************Dump Dictionaries to Dump Sheet ******************************
Public Sub Dump_All_Dictionaries()
    Dim wsDump As Worksheet
    Dim r As Long, c As Long
    Dim key As Variant, subDict As Scripting.Dictionary
    Dim value As Variant

    ' Create or clear the dump sheet
    On Error Resume Next
    Set wsDump = ThisWorkbook.Sheets("dictionary_Dumps")
    If wsDump Is Nothing Then
        Set wsDump = ThisWorkbook.Sheets.Add
        wsDump.Name = "dictionary_Dumps"
    Else
        wsDump.Cells.Clear
    End If
    On Error GoTo 0

    c = 1  ' Start in column A

    ' Dump dict_Design_Options
    If Not dict_Design_Options Is Nothing Then
        wsDump.Cells(1, c).Resize(1, 5).value = Array("Design Option", "Design Abbr", "Price", "Equipment Count", "Equipment List")
        r = 2
        For Each key In dict_Design_Options.Keys
            Set subDict = dict_Design_Options(key)
            wsDump.Cells(r, c).value = key
            wsDump.Cells(r, c + 1).value = subDict("Abbr")
            wsDump.Cells(r, c + 2).value = subDict("Price")
            wsDump.Cells(r, c + 3).value = UBound(subDict("Equipment")) - LBound(subDict("Equipment")) + 1
            wsDump.Cells(r, c + 4).value = Join(subDict("Equipment"), ", ")
            r = r + 1
        Next key
        c = c + 6  ' Leave a gap column
    End If

    ' Dump dict_Fabrics
    If Not dict_Fabrics Is Nothing Then
        wsDump.Cells(1, c).Resize(1, 3).value = Array("Fabric Type", "Fabric Abbr", "Cost Per Sq Inch")
        r = 2
        For Each key In dict_Fabrics.Keys
            Set subDict = dict_Fabrics(key)
            wsDump.Cells(r, c).value = key
            wsDump.Cells(r, c + 1).value = subDict("Abbr")
            wsDump.Cells(r, c + 2).value = subDict("CostPerSqInch")
            r = r + 1
        Next key
        c = c + 4
    End If

    ' Dump dict_Color_Names
    If Not dict_Color_Names Is Nothing Then
        wsDump.Cells(1, c).Resize(1, 4).value = Array("Color Abbr", "Color Name", "Color Map", "Available Fabrics")
        r = 2
        For Each key In dict_Color_Names.Keys
            Set subDict = dict_Color_Names(key)
            wsDump.Cells(r, c).value = key
            wsDump.Cells(r, c + 1).value = subDict("Name")
            wsDump.Cells(r, c + 2).value = subDict("Map")
            wsDump.Cells(r, c + 3).value = SafeJoin(subDict("Available"))

            r = r + 1
        Next key
        c = c + 5
    End If

    ' Dump dict_ShippingRates
    If Not dict_Shipping Is Nothing Then
        wsDump.Cells(1, c).Resize(1, 2).value = Array("Weight", "Shipping Cost")
        r = 2
        For Each key In dict_Shipping.Keys
            wsDump.Cells(r, c).value = key
            wsDump.Cells(r, c + 1).value = dict_Shipping(key)
            r = r + 1
        Next key
        c = c + 3
    End If

    ' Dump dict_Miscellaneous
    If Not dict_Miscellaneous Is Nothing Then
        wsDump.Cells(1, c).Resize(1, 2).value = Array("Field Name", "Value")
        r = 2
        For Each key In dict_Miscellaneous.Keys
            value = dict_Miscellaneous(key)
            wsDump.Cells(r, c).value = key
            If IsArray(value) Then
                wsDump.Cells(r, c + 1).value = Join(value, ", ")
            Else
                wsDump.Cells(r, c + 1).value = value
            End If
            r = r + 1
        Next key
    End If
    ' ? Hide the dump sheet after writing
    wsDump.Visible = xlSheetHidden
    'MsgBox "All dictionaries dumped to '" & wsDump.Name & "'.", vbInformation
End Sub
'*********************************END Dump Dictionaries to Dump Sheet ******************************
Public Sub Unhide_Dictionary_Sheets()
    Dim sheetNames As Variant, i As Long
    sheetNames = Array("var_Design_Options", "var_Fabric_Types", "var_Colors", "var_Shipping", "var_Miscellaneous")

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        ThisWorkbook.Sheets(sheetNames(i)).Visible = xlSheetVisible
        On Error GoTo 0
    Next i
End Sub
Public Sub Hide_Dictionary_Sheets()
    Dim sheetNames As Variant, i As Long
    sheetNames = Array("var_Design_Options", "var_Fabric_Types", "var_Colors", "var_Shipping", "var_Miscellaneous")

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        ThisWorkbook.Sheets(sheetNames(i)).Visible = xlSheetVeryHidden
        On Error GoTo 0
    Next i
End Sub
Public Sub Developer_Unhide_DictionarySheets()
Attribute Developer_Unhide_DictionarySheets.VB_ProcData.VB_Invoke_Func = "U\n14"
    Dim sheetNames As Variant, i As Long
    sheetNames = Array("var_Design_Options", "var_Fabric_Types", "var_Colors", "var_Shipping", "var_Miscellaneous", "dictionary_Dumps")

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        ThisWorkbook.Sheets(sheetNames(i)).Visible = xlSheetVisible
        On Error GoTo 0
    Next i

    MsgBox "All dictionary sheets are now visible.", vbInformation
End Sub
Public Sub Developer_Hide_DictionarySheets()
Attribute Developer_Hide_DictionarySheets.VB_ProcData.VB_Invoke_Func = "H\n14"
    Dim sheetNames As Variant, i As Long
    sheetNames = Array("var_Design_Options", "var_Fabric_Types", "var_Colors", "var_Shipping", "var_Miscellaneous", "dictionary_Dumps")

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        ThisWorkbook.Sheets(sheetNames(i)).Visible = xlSheetVeryHidden
        On Error GoTo 0
    Next i

    MsgBox "All dictionary sheets are now hidden.", vbInformation
End Sub

Public Sub filter_Colors_By_Fabric(ByRef frm As Object, ByVal fabricAbbr As String)
    frm.lst_Fabric_Color_Names.Clear

    Dim colorKey As Variant
    For Each colorKey In dict_Color_Names.Keys
        Dim subDict As Scripting.Dictionary
        Set subDict = dict_Color_Names(colorKey)

        If subDict.Exists("Available") And subDict.Exists("Name") Then
            If IsInArray(fabricAbbr, subDict("Available")) Then
                frm.lst_Fabric_Color_Names.AddItem subDict("Name") & " (" & colorKey & ")"
                Debug.Print "? Added: " & subDict("Name") & " (" & colorKey & ")"
            Else
                Debug.Print "?? Skipped: " & subDict("Name") & " (" & colorKey & ")"
            End If
        Else
            Debug.Print "? Missing fields for colorKey: " & colorKey
        End If
    Next colorKey
End Sub

Function IsInCommaList(ByVal valToCheck As String, ByVal csv As String) As Boolean
    Dim parts() As String
    parts = Split(csv, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Trim(UCase(parts(i))) = UCase(valToCheck) Then
            IsInCommaList = True
            Exit Function
        End If
    Next i

    IsInCommaList = False
End Function
Public Sub load_All_Colors(ByRef frm As Object)
    'RUNNING 07-21-2025
    frm.lst_Fabric_Colors.Clear

    Debug.Print "?? dict_Color_Names count: " & dict_Color_Names.Count

    Dim colorList As Collection
    Set colorList = New Collection

    Dim colorKey As Variant
    For Each colorKey In dict_Color_Names.Keys
        Debug.Print "?? Checking colorKey: " & colorKey

        Dim subDict As Scripting.Dictionary
        Set subDict = dict_Color_Names(colorKey)

        If subDict.Exists("My Color Name") Then
            Dim colorName As String
            colorName = Trim(subDict("My Color Name"))

            If Len(colorName) > 0 And UCase(colorName) <> "SKIP" Then
                colorList.Add colorName & " (" & colorKey & ")"
                Debug.Print "? Queued Color: " & colorName & " (" & colorKey & ")"
            Else
                Debug.Print "?? Skipped empty or SKIP color: " & colorKey
            End If
        Else
            Debug.Print "? 'My Color Name' key missing for: " & colorKey
        End If
    Next colorKey

    ' Convert and sort the collection
    Dim sortedColors As Variant
    sortedColors = SortVariantArray(CollectionToArray(colorList))

    ' Populate the list box
    Dim i As Long
    For i = LBound(sortedColors) To UBound(sortedColors)
        frm.lst_Fabric_Colors.AddItem sortedColors(i)
        Debug.Print "? Added to list: " & sortedColors(i)
    Next i
End Sub







Public Function SortVariantArray(ByVal inputArray As Variant) As Variant
    Dim tempArray() As String
    Dim i As Long

    ' Copy input to temp array
    ReDim tempArray(LBound(inputArray) To UBound(inputArray))
    For i = LBound(inputArray) To UBound(inputArray)
        tempArray(i) = inputArray(i)
    Next i

    ' Sort using VBA's built-in array sort
    Dim j As Long, temp As String
    For i = LBound(tempArray) To UBound(tempArray) - 1
        For j = i + 1 To UBound(tempArray)
            If StrComp(tempArray(i), tempArray(j), vbTextCompare) > 0 Then
                temp = tempArray(i)
                tempArray(i) = tempArray(j)
                tempArray(j) = temp
            End If
        Next j
    Next i

    SortVariantArray = tempArray
End Function
Public Function CollectionToArray(col_Platform_List) As Variant

    Dim arr() As Variant

    If col_Platform_List.Count = 0 Then
        Debug.Print "?? Collection is empty in CollectionToArray"
        CollectionToArray = Array()
        Exit Function
    End If

    ReDim arr(0 To col_Platform_List.Count - 1)

    Dim i As Long
    For i = 1 To col_Platform_List.Count
        arr(i - 1) = col_Platform_List(i)
        Debug.Print "?? arr[" & (i - 1) & "] = " & TypeName(arr(i - 1)) & ": " & arr(i - 1)
    Next i

    CollectionToArray = arr
End Function

Public Function IsInArray(valueToFind As String, arr As Variant) As Boolean
    Dim i As Long

    If Len(Trim(valueToFind)) = 0 Then
        Debug.Print "?? IsInArray skipped — valueToFind is empty"
        Exit Function
    End If

    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            Debug.Print "?? Comparing: [" & UCase(Trim(arr(i))) & "] vs [" & UCase(Trim(valueToFind)) & "]"
            If Trim(UCase(arr(i))) = Trim(UCase(valueToFind)) Then
                Debug.Print "? Match found for [" & valueToFind & "]"
                IsInArray = True
                Exit Function
            End If
        Next i
    Else
        Debug.Print "?? IsInArray skipped — input is not an array"
    End If

    IsInArray = False
End Function




Public Function SafeJoin(val As Variant, Optional delimiter As String = ", ") As String
    Select Case TypeName(val)
        Case "String"
            SafeJoin = val
        Case "Variant()", "String()", "Array"
            SafeJoin = Join(val, delimiter)
        Case Else
            SafeJoin = "[Invalid Type: " & TypeName(val) & "]"
    End Select
End Function




    
    Public Sub dump_Fabric_Display_Map()
    Dim key As Variant
    Debug.Print "?? Dumping fabric_Display_Map..."
    For Each key In fabric_Display_Map.Keys
        Debug.Print "?? " & key & " ? " & fabric_Display_Map(key)
    Next key
    Debug.Print "? End of map"
End Sub


    
    




    
    
    
  


