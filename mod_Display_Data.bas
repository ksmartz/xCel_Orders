Attribute VB_Name = "mod_Display_Data"
Public Sub Populate_InputSheet_FromModelDictionary()
    Dim wsInput As Worksheet: Set wsInput = ThisWorkbook.Sheets("Input")
    Dim r As Long: r = 2
    Dim modelKey As Variant
    Dim modelDict As Scripting.Dictionary

    Dim totalSqInch As Double
    Dim fabricWeights As Variant
    Dim fabricCosts As Variant
    Dim totalCosts As Variant
    Dim retailPrices As Variant

    wsInput.Range("A2:Z1000").ClearContents  ' Reset the sheet

    For Each modelKey In dict_Models.Keys
        Set modelDict = dict_Models(modelKey)

        ' === Base Info ===
        wsInput.Cells(r, 1).value = str_Manufacturer_Name
        wsInput.Cells(r, 2).value = str_Series_Name
        wsInput.Cells(r, 3).value = modelDict("Model Name")
        wsInput.Cells(r, 4).value = str_Equipment_Type
        wsInput.Cells(r, 5).value = modelDict("Width")
        wsInput.Cells(r, 6).value = modelDict("Depth")
        wsInput.Cells(r, 7).value = modelDict("Height")

        ' === Fabric Area ===
        totalSqInch = GetAdjustedFabricArea(modelDict)
        wsInput.Cells(r, 8).value = totalSqInch

        ' === Weight per Fabric Type ===
        fabricWeights = GetFabricWeights(totalSqInch)
        wsInput.Cells(r, 9).Resize(1, 4).value = fabricWeights  ' Columns I–L

        ' === Material Cost per Fabric ===
       fabricCosts = GetFabricCosts(totalSqInch)
'        wsInput.Cells(r, 13).Resize(1, 4).value = fabricCosts  ' Columns M–P

        ' === Final Cost (includes everything except profit uplift) ===
        totalCosts = GetTotalCoverCosts(fabricWeights, fabricCosts)
        wsInput.Cells(r, 13).Resize(1, 4).value = totalCosts  ' Columns Q–T

        ' === Final Retail Price (Profit Adjusted Only) ===
        retailPrices = GetRetailPrices(totalCosts)
        wsInput.Cells(r, 17).Resize(1, 4).value = retailPrices  ' Columns U–X

        r = r + 1
    Next modelKey

    Debug.Print "? Sheet populated with " & dict_Models.Count & " models."
End Sub


Public Sub WriteModelDimensionsToOrderSheet(ByRef frm As Object)
    Dim modelName As String
    modelName = frm.lst_Models.value
    Debug.Print "?? Selected model: [" & modelName & "]"

    ' ?? Validate model dictionary
    If dict_Models Is Nothing Then
        Debug.Print "? dict_Models is not initialized."
        Exit Sub
    End If

    If Not dict_Models.Exists(modelName) Then
        Debug.Print "? Model not found in dict_Models: [" & modelName & "]"
        Exit Sub
    End If

    Dim modelDict As Scripting.Dictionary
    Set modelDict = dict_Models(modelName)

    ' ?? Get selected form block index (opt_OrderX)
    Dim blockIndex As Long
    blockIndex = get_Selected_Block_Index(frm)

    If blockIndex = -1 Then
        Debug.Print "? No order block selected via opt_OrderX."
        Exit Sub
    End If

    ' ?? Get cached anchor blocks from sheet
    Dim allBlocks() As order_Information_Block
    allBlocks = get_Order_Type_Block_List()

    If blockIndex > UBound(allBlocks) Then
        Debug.Print "? Block index [" & blockIndex & "] exceeds available blocks [" & UBound(allBlocks) & "]"
        Exit Sub
    End If

    Dim anchorRow As Long
    anchorRow = CLng(allBlocks(blockIndex).str_Anchor_Row)

    Debug.Print "?? Writing to anchor row: " & anchorRow
    Debug.Print "?? Dimensions ? W:" & modelDict("Width") & " D:" & modelDict("Depth") & _
                " H:" & modelDict("Height") & " OptDepth:" & modelDict("Opt. Depth")
                
 

    ' ?? Write values to sheet
    With ws_Orders
    .Cells(anchorRow + 1, 6).value = modelName                        ' Column F: Model Name
    .Cells(anchorRow + 4, 1).value = modelDict("Width")              ' Column A: Width
    .Cells(anchorRow + 4, 2).value = modelDict("Depth")              ' Column B: Depth
    .Cells(anchorRow + 4, 3).value = modelDict("Height")             ' Column C: Height

    If modelDict.Exists("Opt. Depth") Then
        .Cells(anchorRow + 4, 4).value = modelDict("Opt. Depth")     ' Column D: Optional Depth
    Else
        Debug.Print "?? Optional Depth missing for model [" & modelName & "]"
        .Cells(anchorRow + 4, 4).value = ""                          ' Optional: clear or placeholder
    End If

    If modelDict.Exists("Angle Type") Then
        .Cells(anchorRow + 4, 5).value = modelDict("Angle Type")     ' Column E: Angle Type
    Else
        Debug.Print "?? Angle Type missing for model [" & modelName & "]"
        .Cells(anchorRow + 4, 5).value = ""
    End If

    If modelDict.Exists("Opt. Height") Then
        .Cells(anchorRow + 4, 6).value = modelDict("Opt. Height")    ' Column F: Optional Height
    Else
        Debug.Print "?? Optional Height missing for model [" & modelName & "]"
        .Cells(anchorRow + 4, 6).value = ""
    End If
    
    If str_Equipment_Type = "Guitar Amp" Then
        Call WriteAmpHandleStringToOrderBlock(modelName, modelDict, anchorRow)
    Else
    End If
    Call Dispatch_1_Piece_Calculation(modelName, modelDict, anchorRow)

End With


    Debug.Print "? Dimensions written for model [" & modelName & "] into block index [" & blockIndex & "]"
End Sub



Public Sub WriteAmpHandleStringToOrderBlock(ByVal modelName As String, ByRef modelDict As Scripting.Dictionary, ByVal anchorRow As Long)
    Dim ahLabel As String, ahLength As String, ahWidth As String
    Dim ampHandleText As String

    ahLabel = IIf(modelDict.Exists("AH: Location"), modelDict("AH: Location"), "")
    ahLength = IIf(modelDict.Exists("TAH/SAH: Length/Height"), modelDict("TAH/SAH: Length/Height"), "")
    ahWidth = IIf(modelDict.Exists("TAH/SAH: Width"), modelDict("TAH/SAH: Width"), "")

    If Len(Trim(ahLabel)) > 0 Or Len(Trim(ahLength)) > 0 Or Len(Trim(ahWidth)) > 0 Then
        ampHandleText = ahLabel & ":" & vbCrLf & ahLength & """L x " & ahWidth & """W"
    Else
        ampHandleText = ""
        Debug.Print "?? Amp handle string not generated for [" & modelName & "]"
    End If

    With ws_Orders.Cells(anchorRow + 6, 5)
        .value = ampHandleText
        .WrapText = True  ' ? Ensure the cell supports multiline display
    End With

    Debug.Print "?? Amp handle (wrapped) ? [" & Replace(ampHandleText, vbCrLf, " ? ") & "] written to Row: " & (anchorRow + 6) & ", Col: 5"
End Sub








