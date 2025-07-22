Attribute VB_Name = "Mod_Calculations"
Public Function RoundUpToEighth(inValue As Double) As Double
    RoundUpToEighth = Application.WorksheetFunction.Ceiling(inValue, 0.125)
End Function
Public Function RoundUpTo95(val As Variant) As Double
    Dim wholePart As Long
    wholePart = Int(val)
    
    If val <= wholePart + 0.95 Then
        RoundUpTo95 = wholePart + 0.95
    Else
        RoundUpTo95 = wholePart + 1.95
    End If
End Function


Public Function GetShippingCostFromWeight(wt As Variant) As Double
    Dim roundedWeight As Long
    Dim k As Variant

    ' Round up to nearest whole ounce
    roundedWeight = Application.WorksheetFunction.Ceiling_Math(wt, 1)

    ' Check for direct match first
'    If dict_Shipping.Exists(CStr(roundedWeight)) Then
'        GetShippingCostFromWeight = dict_Shipping(CStr(roundedWeight))
'        Exit Function
'    End If
    
    
    
    If dict_Shipping.Exists(roundedWeight) Then
    GetShippingCostFromWeight = dict_Shipping(roundedWeight)
    Exit Function
End If


    ' Handle custom range-style keys
    For Each k In dict_Shipping.Keys
        Select Case k
            Case ">16 <32"
                If roundedWeight > 16 And roundedWeight < 32 Then
                    GetShippingCostFromWeight = dict_Shipping(k)
                    Exit Function
                End If
            Case ">=32 <48"
                If roundedWeight >= 32 And roundedWeight < 48 Then
                    GetShippingCostFromWeight = dict_Shipping(k)
                    Exit Function
                End If
            Case ">=48 <60"
                If roundedWeight >= 48 And roundedWeight < 60 Then
                    GetShippingCostFromWeight = dict_Shipping(k)
                    Exit Function
                End If
            Case ">60"
                If roundedWeight > 60 Then
                    GetShippingCostFromWeight = dict_Shipping(k)
                    Exit Function
                End If
        End Select
    Next k

    ' Fallback: no match found
    GetShippingCostFromWeight = 0
End Function







Public Function GetAdjustedFabricArea(modelDict As Scripting.Dictionary) As Double
    Dim w As Double, D As Double, H As Double
    Dim faceArea As Double, sideArea As Double, topArea As Double
    Dim baseArea As Double, paddedArea As Double

    ' Round dimensions to nearest 1/8 inch
    If modelDict.Exists("Width") Then w = RoundUpToEighth(modelDict("Width"))
    If modelDict.Exists("Depth") Then D = RoundUpToEighth(modelDict("Depth"))
    If modelDict.Exists("Height") Then H = RoundUpToEighth(modelDict("Height"))

    ' Apply seam allowances
    w = w + 1
    D = D + 1
    H = H + 1

    ' Calculate total base area
    faceArea = 2 * w * H
    sideArea = 2 * D * H
    topArea = w * D
    baseArea = faceArea + sideArea + topArea

    ' Add 5% waste allowance
    paddedArea = baseArea * 1.05

    ' Return as rounded-up square inches
    GetAdjustedFabricArea = Application.WorksheetFunction.Ceiling_Math(paddedArea, 1)
End Function

Public Function GetFabricWeights(totalSqInch As Double) As Variant
    Dim fabricTypes As Variant: fabricTypes = Array("C", "CG", "L", "LG")
    Dim weights(0 To 3) As Double
    Dim i As Long, abbr As String
    Dim ozPerSqInch As Double

    For i = 0 To UBound(fabricTypes)
        abbr = fabricTypes(i)
        If dict_Fabrics.Exists(abbr) Then
            If dict_Fabrics(abbr).Exists("Ounces per Square Inch") Then
                ozPerSqInch = val(dict_Fabrics(abbr)("Ounces per Square Inch"))
                If totalSqInch > 0 And ozPerSqInch > 0 Then
                
                Debug.Print abbr & ": Area=" & totalSqInch & ", OZ/in²=" & ozPerSqInch & ", Weight=" & totalSqInch * ozPerSqInch

                    weights(i) = Application.WorksheetFunction.Ceiling_Math( _
                                    totalSqInch * ozPerSqInch, 1)
                Else
                    weights(i) = 0  ' fallback
                End If
            End If
        End If
    Next i

    GetFabricWeights = weights
End Function
'Public Function GetTotalCoverCosts(fabricCosts As Variant) As Variant
'    Dim costs(0 To 3) As Double
'    Dim laborHours As Double
'    Dim hourlyRate As Double, bagExpense As Double
'    Dim fabricTypes As Variant: fabricTypes = Array("C", "CG", "L", "LG")
'    Dim i As Long, abbr As String
'
'    ' Fixed cost components
'    hourlyRate = Val(dict_Miscellaneous("Hourly Rate"))
'    bagExpense = Val(dict_Miscellaneous("Bag Expense"))
'
'    For i = 0 To UBound(fabricTypes)
'        abbr = fabricTypes(i)
'
'        Select Case abbr
'            Case "CG", "LG": laborHours = Val(dict_Miscellaneous("CG,LG Labor"))
'            Case "C", "L":   laborHours = Val(dict_Miscellaneous("C, L Labor"))
'        End Select
'
'        costs(i) = Round(fabricCosts(i) + (hourlyRate * laborHours) + bagExpense, 2)
'    Next i
'
'    GetTotalCoverCosts = costs
'End Function


Public Function GetFabricCosts(totalSqInch As Double) As Variant
    Dim fabricTypes As Variant: fabricTypes = Array("C", "CG", "L", "LG")
    Dim costs(0 To 3) As Double
    Dim i As Long, abbr As String
    Dim costPerSqIn As Double

    For i = 0 To UBound(fabricTypes)
        abbr = fabricTypes(i)
        If dict_Fabrics.Exists(abbr) Then
            If dict_Fabrics(abbr).Exists("Cost per Square Inch") Then
                costPerSqIn = val(dict_Fabrics(abbr)("Cost per Square Inch"))
                If totalSqInch > 0 And costPerSqIn > 0 Then
                Debug.Print "abbr=" & abbr & ", Rate=" & dict_Fabrics(abbr)("Cost per Square Inch") & ", Parsed=" & val(dict_Fabrics(abbr)("Cost per Square Inch"))

                    costs(i) = Round(totalSqInch * costPerSqIn, 2)
                End If
            End If
        End If
    Next i
    Debug.Print "abbr=" & abbr & ", Rate=" & dict_Fabrics(abbr)("Cost per Square Inch") & ", Parsed=" & val(dict_Fabrics(abbr)("Cost per Square Inch"))

    GetFabricCosts = costs
End Function


'************************Calculate Fabric Cost***************************
Public Function calculate_Fabric_Weight_And_Cost()
    Dim db_Width As Double
    Dim db_Yards As Double
    Dim db_Cost_Per_Yard As Double
    Dim db_Shipping_Cost As Double
    Dim db_Fabric_Subtotal As Double
    Dim db_Total_Cost As Double
    Dim db_Total_Square_Inches As Double
    Dim db_Cost_Per_Square_Inch As Double

    ' Validate inputs
    If txt_Fabric_Width.value = "" Or txt_Amount_Of_Linear_Yards.value = "" Or txt_Cost_Per_Linear_Yard.value = "" Or txt_Shipping_Cost.value = "" Then
        MsgBox "Please fill in all fields.", vbExclamation
        Exit Sub
    End If

    ' Convert inputs
    db_Width = val(txt_Fabric_Width.value)
    db_Yards = val(txt_Amount_Of_Linear_Yards.value)
    db_Cost_Per_Yard = val(txt_Cost_Per_Linear_Yard.value)
    db_Shipping_Cost = val(txt_Shipping_Cost.value)

    ' Core calculations
    db_Fabric_Subtotal = db_Cost_Per_Yard * db_Yards
    db_Total_Cost = db_Fabric_Subtotal + db_Shipping_Cost
    db_Total_Square_Inches = db_Width * (db_Yards * 36)
    db_Cost_Per_Square_Inch = db_Total_Cost / db_Total_Square_Inches

    ' Output results
    lbl_Total_Square_Inches_Holder.Caption = Format(db_Total_Square_Inches, "#,##0") & " sq in"
    lbl_Subtotal_Fabric_Cost_Holder.Caption = "$" & Format(db_Fabric_Subtotal, "0.00")
    lbl_Total_Fabric_Cost_Holder.Caption = "$" & Format(db_Total_Cost, "0.00")
    lbl_Cost_Per_Square_Inch_Holder.Caption = "$" & Format(db_Cost_Per_Square_Inch, "0.000000")
End Function

'
'Public Function GetMPCosts(totalCosts As Variant, marketplace As String) As Variant
'    Dim fabricTypes As Variant: fabricTypes = Array("C", "CG", "L", "LG")
'    Dim finalCosts(0 To 3) As Double
'    Dim i As Long, abbr As String
'
'    Dim profitAdj As Double
'    Dim mpFee As Double
'    Dim rawCost As Double, retailPrice As Double
'
'    If dict_MarketPlace_Specifics.Exists(marketplace) Then
'        mpFee = Val(dict_MarketPlace_Specifics(marketplace)("Sales Percentage")) / 100
'    Else
'        mpFee = 0  ' default if marketplace not found
'    End If
'
'    For i = 0 To UBound(fabricTypes)
'        abbr = fabricTypes(i)
'
'        If dict_Fabrics.Exists(abbr) Then
'            If dict_Fabrics(abbr).Exists("Profit Adjustment") Then
'                profitAdj = Val(dict_Fabrics(abbr)("Profit Adjustment"))
'
'                If totalCosts(i) > 0 Then
'                    ' Reverse out profit uplift
'                    rawCost = totalCosts(i) / (1 + profitAdj / 100)
'                    ' Add marketplace fee
'                    retailPrice = Round(rawCost + (rawCost * mpFee), 2)
'                    finalCosts(i) = retailPrice
'                End If
'            End If
'        End If
'    Next i
'
'    GetMPCosts = finalCosts
'End Function

Public Function GetTotalCoverCosts(fabricWeights As Variant, fabricCosts As Variant) As Variant
    Dim costs(0 To 3) As Double
    Dim i As Long, abbr As String
    Dim hourlyRate As Double, bagExpense As Double, laborHours As Double
    Dim profitAdj As Double, mpFeePct As Double
    Dim fabricTypes As Variant: fabricTypes = Array("C", "CG", "L", "LG")

    Dim baseCost As Double
    Dim adjustedCost As Double
    Dim marketplaceFee As Double
    Dim shippingCost As Double

    hourlyRate = val(dict_Miscellaneous("Hourly Rate"))
    bagExpense = val(dict_Miscellaneous("Bag Expense"))
    mpFeePct = val(dict_MarketPlace_Specifics("Amazon")("Sales Percentage")) / 100

    For i = 0 To 3
        abbr = fabricTypes(i)

        If dict_Fabrics(abbr).Exists("Profit Adjustment") Then
            profitAdj = val(dict_Fabrics(abbr)("Profit Adjustment"))

            Select Case abbr
                Case "CG", "LG": laborHours = val(dict_Miscellaneous("CG,LG Labor"))
                Case "C", "L":   laborHours = val(dict_Miscellaneous("C, L Labor"))
            End Select
            Debug.Print "Weight[" & i & "] = " & fabricWeights(i)
Debug.Print "Shipping cost = " & GetShippingCostFromWeight(fabricWeights(i))

            shippingCost = GetShippingCostFromWeight(fabricWeights(i))

            baseCost = fabricCosts(i) + (hourlyRate * laborHours) + bagExpense + shippingCost
            adjustedCost = baseCost * (1 + profitAdj / 100)
            marketplaceFee = adjustedCost * mpFeePct

            costs(i) = Round(baseCost + marketplaceFee, 2)
        End If
    Next i

    GetTotalCoverCosts = costs
End Function

Public Function GetRetailPrices(totalCosts As Variant) As Variant
    Dim retail(0 To 3) As Double
    Dim i As Long, abbr As String
    Dim profitAdj As Double
    Dim fabricTypes As Variant: fabricTypes = Array("C", "CG", "L", "LG")

    For i = 0 To 3
        abbr = fabricTypes(i)

        If dict_Fabrics.Exists(abbr) Then
            If dict_Fabrics(abbr).Exists("Profit Adjustment") Then
                profitAdj = val(dict_Fabrics(abbr)("Profit Adjustment"))
                retail(i) = RoundUpTo95(totalCosts(i) + profitAdj)
            Else
                retail(i) = RoundUpTo95(totalCosts(i))  ' no uplift defined
            End If
        Else
            retail(i) = RoundUpTo95(totalCosts(i))  ' fabric type missing
        End If
    Next i

    GetRetailPrices = retail
End Function

Public Sub Dispatch_1_Piece_Calculation(ByVal modelName As String, ByRef modelDict As Scripting.Dictionary, ByVal anchorRow As Long)
    Dim equipmentType As String, angleType As String
    equipmentType = IIf(modelDict.Exists("Equipment Type"), modelDict("Equipment Type"), "")
    angleType = IIf(modelDict.Exists("Angle Type"), modelDict("Angle Type"), "")

    Select Case Trim(equipmentType)
        Case "Music Keyboard"
            Call Calculate_1_Piece_Keyboard(modelName, modelDict, anchorRow)

        Case "Guitar Amp"
            Call Calculate_1_Piece_GuitarAmp(modelName, modelDict, anchorRow, angleType)

        Case Else
            Debug.Print "?? Unknown Equipment Type [" & equipmentType & "] for model [" & modelName & "]; no calculation applied."
    End Select
End Sub


Public Sub Calculate_1_Piece_Keyboard(ByVal modelName As String, ByRef modelDict As Scripting.Dictionary, ByVal anchorRow As Long)
    Dim widthVal As Double, depthVal As Double, heightVal As Double
    Dim db_Model_Width_1Piece As Double, db_Model_Height_1Piece As Double

    ' ?? Defensive extraction
    widthVal = IIf(modelDict.Exists("Width"), modelDict("Width"), 0)
    depthVal = IIf(modelDict.Exists("Depth"), modelDict("Depth"), 0)
    heightVal = IIf(modelDict.Exists("Height"), modelDict("Height"), 0)

    ' ?? Compute 1-Piece dimensions
    db_Model_Width_1Piece = (widthVal + 1.25) + (heightVal + 1) + heightVal
    db_Model_Width_1Piece = Round(db_Model_Width_1Piece * 8, 0) / 8

    db_Model_Height_1Piece = (depthVal + 1.25) + heightVal + (heightVal + 0.5)
    db_Model_Height_1Piece = Round(db_Model_Height_1Piece * 8, 0) / 8

    ' ?? Write to Orders sheet
    With ws_Orders
        .Cells(anchorRow + 6, 2).value = db_Model_Width_1Piece     ' Row 7, Column B
        .Cells(anchorRow + 6, 4).value = db_Model_Height_1Piece    ' Row 7, Column D
    End With

    ' ?? Debug trace
    Debug.Print "?? Calculated 1-Piece dimensions for [" & modelName & "] ?"
    Debug.Print "     Width:  " & db_Model_Width_1Piece & " ? Row " & (anchorRow + 6) & ", Col 2"
    Debug.Print "     Height: " & db_Model_Height_1Piece & " ? Row " & (anchorRow + 6) & ", Col 4"
End Sub

'Sub calculate_1_Piece()
'
'    Select Case str_Equipment_Type
'        Case Is = "Guitar Amp"
'            Select Case str_Angle_Type
'                Case Is = "Full-Angle", "Full-Curve", "Mid-Curve"
'                    'db_Model_Width_1Piece = (Height + 1.25) + Depth_Optional + db_Angle_Line_Length + 1.25
'                Case Is = "Top-Angle"
'                   ' db_Model_Width_1Piece = (Height + 1.25) + (db_Depth_Optional) + (db_Angle_Line_Length + (db_Height_Optional - 1.25))
'                Case Is = "Mid-Angle"
'                    'db_Model_Width_1Piece = Round(((db_Model_Height) + (db_Depth_Optional + 0.75) + db_Angle_Line_Length + db_Height_Optional) * 8, 0) / 8
'                Case Is = ""
'                    'db_Model_Width_1Piece = ((db_O_Height + 1.25) * 2) + db_O_Depth
'            End Select
'
'            'db_Model_Width_1Piece = Round(db_Model_Width_1Piece * 8, 0) / 8
'            'db_Model_Height_1Piece = Round((db_Model_Width + 0.75) * 8, 0) / 8
'
'
'
'        Case Is = "Music Keyboard"
'            db_Model_Width_1Piece = (Width + 1.25) + (Height + 1) + (Height)
'            db_Model_Width_1Piece = Round((db_Model_Width_1Piece) * 8, 0) / 8
'
'            db_Model_Height_1Piece = ((Depth + 1.25) + (Height) + (Height + 0.5))
'            db_Model_Height_1Piece = Round((db_Model_Height_1Piece) * 8, 0) / 8
'    End Select
'End Sub

