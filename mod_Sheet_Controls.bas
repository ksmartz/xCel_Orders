Attribute VB_Name = "mod_Sheet_Controls"
Public Sub handle_chk_Hidden_Dictionary_Sheets()
    Dim sheetNames As Variant, i As Long
    sheetNames = Array("var_Design_Options", "var_Fabric_Types", "var_Colors", "var_Shipping", "var_Miscellaneous")

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        ThisWorkbook.Sheets(sheetNames(i)).Visible = xlSheetVeryHidden
        On Error GoTo 0
    Next i
End Sub
Public Sub handle_Btn_New_Listings()
frm_New_Listings.Show

End Sub
Public Sub handle_Btn_Input_Sheet()
    Worksheets("Input").Activate
End Sub

Public Sub handle_Btn_Order_Sheet()
Worksheets("Orders").Activate

End Sub




Public Sub handle_Btn_Brand_Names()


End Sub

Public Sub handle_Btn_Template_Sheet()


    Dim tbl As ListObject
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Assumes there's only one table on the sheet
    If ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
        tbl.Name = "tbl_" & ws.Name & "_Series_Name"
    End If





End Sub
Public Sub handle_Btn_Calculate_Fabric_Cost()


End Sub

Public Sub handle_Btn_Unprotect_Sheet()


End Sub
Public Sub handle_Btn_Protect_Sheet()


End Sub
Public Sub handle_Btn_Protect_Workbook()


End Sub
Public Sub handle_Btn_Unprotect_Workbook()


End Sub
Public Sub handle_Toggle_Autolock()


End Sub
Public Sub handle_Btn_Home()
    Application.GoTo reference:=Range("A1"), Scroll:=True
End Sub
Public Sub handle_Btn_Search_Worksheet()


End Sub
Public Sub handle_txt_Search_Work_Sheet()


End Sub















Public Sub SelectListItem(ctrl As Object, value As Variant)
    Dim i As Long
    For i = 0 To ctrl.ListCount - 1
        If ctrl.List(i) = value Then
            ctrl.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub




Public Function GetSelectedOption(frm As Object) As String
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.value = True Then
                GetSelectedOption = ctrl.Name
                Exit Function
            End If
        End If
    Next ctrl
    GetSelectedOption = ""
End Function



Private Sub ResetFormControls(frm As Object)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.value = ""
            Case "ComboBox", "ListBox"
                ctrl.ListIndex = -1
            Case "CheckBox", "OptionButton"
                ctrl.value = False
        End Select
    Next ctrl
End Sub

Private Sub ClearOption2Data()
    With Sheets("Orders")
        .Range("B10").ClearContents
        .Range("B11:B15").ClearContents
    End With
End Sub

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




Public Sub ShowOrderReviewForm(data As Object, frm As Object, targetCell As Range)
    Dim fReview As frm_Order_Review
    Set fReview = frm_Order_Review
    
    
   
    

    ' ?? Assign references to calling form and order block
    Set fReview.frmTarget = frm               ' ? Calling form (e.g. frm_New_Orders)
    Set fReview.targetCell = targetCell

    ' ?? Debug validation
    Debug.Print "?? TargetCell Assigned to Review Form:"
    Debug.Print "  Address: " & fReview.targetCell.Address
    Debug.Print "  Sheet: " & fReview.targetCell.Worksheet.Name
    Debug.Print "  Top Row of Merge: " & fReview.targetCell.MergeArea.Row
    Debug.Print "  Merge Height: " & fReview.targetCell.MergeArea.Rows.Count


    ' ?? Populate review labels
    On Error Resume Next ' skip any missing keys
    With fReview
        .lbl_Name.Caption = data("Name")
        .lbl_Platform_Name.Caption = data("Platform")
        .lbl_Equipment_Type.Caption = data("EquipmentType")
        .lbl_Manufacturer_Name.Caption = data("Manufacturer")
        .lbl_Series_Name.Caption = data("Series")
        .lbl_Model_Name.Caption = data("Model")
        .lbl_Fabric_Type_Name.Caption = data("FabricType")
        .lbl_Color_Name.Caption = data("FabricColor")
       .lbl_Date.Caption = Format(data("Date"), "mm/dd/yyyy")

    End With
    On Error GoTo 0

    ' ?? Launch review form
    fReview.Show
End Sub





