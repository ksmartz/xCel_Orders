VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_New_Orders 
   Caption         =   "Input Customer Orders"
   ClientHeight    =   10370
   ClientLeft      =   -40
   ClientTop       =   -130
   ClientWidth     =   11700
   OleObjectBlob   =   "frm_New_Orders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_New_Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public boo_Frm_Order_Initializing As Boolean
Public boo_Frm_Order_Input As Boolean

Dim boo_click_Triggered_Event As Boolean
Dim rng_Anchor As Range
Dim lng_R As Long
Private fieldBinders() As cls_Order_Form_To_Sheet_Binder
Public boo_Is_Bulk_Clearing As Boolean



'Private testBinder As cls_Order_Form_To_Sheet_Binder




Private Sub cb_MusicRest_Click()
    'cb_MusicRest_Click
End Sub

Private Sub cb_PicPocket_Click()
   ' cb_PicPocket_Click
End Sub

Private Sub cb_PriorityShipping_Click()
    'cb_PriorityShipping_Click
End Sub

Private Sub cb_ZipperHandle_Click()
   ' cb_ZipperHandle_Click
End Sub




Private Sub cb_ZipperHandle_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        
End Sub

Private Sub cmd_Print_Pg_1_Click()
  
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmd_Save_Orders_Click()
 
End Sub

Private Sub lst_Series_Names_AfterUpdate()

End Sub

Private Sub lst_Series_Names_Click()

    
End Sub
Private Sub lst_Series_Names_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  
End Sub

Private Sub lst_Brand_Names_Change()




End Sub



Private Sub btn_Debug_Colors_Click()
      Debug.Print "?? Fabric Diagnostic Triggered"
    Debug.Print "Selected Value: [" & Me.lst_Fabric_Types.value & "]"

    If fabric_Display_Map.Exists(Me.lst_Fabric_Types.value) Then
        Dim abbr As String
        abbr = fabric_Display_Map(Me.lst_Fabric_Types.value)
        Debug.Print "Abbreviation: [" & abbr & "]"
        Call populate_Color_Names(Me, abbr)
    Else
        Debug.Print "?? No abbreviation found for: [" & Me.lst_Fabric_Types.value & "]"
    End If

    Debug.Print "? Final Color List Count: " & Me.lst_Fabric_Colors.ListCount
End Sub

Private Sub btn_Unlock_Sheet_Fields_Click()
    Call unlock_User_Fields
End Sub


Private Sub lst_Fabric_Color_Names_Click()
    
    Set rng_Anchor = get_Order_Anchor_Range(GetSelectedOrderIndex(Me))

    If Not rng_Anchor Is Nothing Then
        rng_Anchor.Offset(1, 1).value = Me.lst_Fabric_Colors.value  ' Column B = Col 2
    End If

    ' ?? Unlock dependent fields
    Me.lst_Manufacturers.Enabled = True
    Me.lst_Series.Enabled = True
    Me.lst_Models.Enabled = True
End Sub





Private Sub lst_Manufacturer_Names_Click()
    Call On_Manufacturer_User_Selection(frm_New_Orders)

    ' ? Live-write to Orders sheet
   
    lng_R = GetOrderStartRow(GetSelectedOrderIndex(Me))
    If lng_R > 0 Then
        ThisWorkbook.Sheets("Orders").Cells(lng_R + 1, 4).value = Me.lst_Manufacturer_Names.value ' Row below start, Column D
    End If
End Sub


Private Sub lst_Model_Names_Click()

    lng_R = get_Order_Anchor_Range(GetSelectedOrderIndex(Me))
    If lng_R > 0 Then
        ThisWorkbook.Sheets("Orders").Cells(lng_R + 1, 6).value = Me.lst_Model_Names.value ' Column F
    End If
    Call On_Model_Name_User_Selection(Me)
End Sub


Private Sub lst_OMP_Names_AfterUpdate()
 
End Sub

Private Sub lst_OMP_Names_Click()
  
End Sub

Private Sub lst_OMP_Names_Enter()
  
End Sub

Private Sub lst_OMP_Names_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub








Private Sub lst_Series_Names_Enter()

End Sub

Private Sub option_Buttons_Handler(ob As MSForms.OptionButton)
   str_Opt_Order = ob.Name
   'MsgBox ob.Name
End Sub

Private Sub lst_Platforms_Names_Click()
    Dim r As Long
    r = GetOrderStartRow(GetSelectedOrderIndex(Me))
    
    If r > 0 Then
        ThisWorkbook.Sheets("Orders").Cells(r, 5).value = Me.lst_Platforms_Names.value ' Column E
    End If

    ' ? Unlock fabric type and options
    Me.lst_Fabric_Type_Names.Enabled = True
   
    Me.fra_Options.Enabled = True
End Sub




Private Sub lst_Series_Name_Click()
    Call mod_Form_Load_Controls.On_Series_Name_User_Selection(Me)

    Dim r As Long
    r = GetOrderStartRow(GetSelectedOrderIndex(Me))
    If r > 0 Then
        ThisWorkbook.Sheets("Orders").Cells(r + 1, 5).value = Me.lst_Series_Name.value ' Column E
    End If


End Sub

Private Sub lst_Fabric_Colors_Click()

End Sub

Private Sub lst_Fabric_Types_Change()
    Call handle_lst_Fabric_Type_Names_Change(frm_New_Orders)
   ' Call filter_Fabric_Colors_By_FabricType(frm_New_Orders)
End Sub

Private Sub lst_Fabric_Types_Click()

End Sub

Private Sub lst_Manufacturers_Click()
    Call mod_Form_Load_Controls.On_Manufacturer_User_Selection(frm_New_Orders)
End Sub

Private Sub lst_Models_Change()
    Call WriteModelDimensionsToOrderSheet(Me)
End Sub

Private Sub lst_Models_Click()

End Sub

Private Sub lst_Series_Click()
    Call On_Series_Name_User_Selection(frm_New_Orders)
End Sub

Private Sub opt_Order1_Click()
    'RUNNING CODE 7-21-2025
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
    Call refresh_Form_From_Selected_Block(frm_New_Orders)
End Sub

Private Sub opt_Order2_Click()
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
    Call clear_Form_Fields(frm_New_Orders)
End Sub

Private Sub opt_Order3_Click()
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
End Sub

Private Sub opt_Order4_Click()
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
End Sub

Private Sub opt_button_Click(opt_Btn As MSForms.OptionButton)


Select Case opt_Btn.Caption
        Case Is = "Order #1"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$2")
        Case Is = "Order #2"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$13")
        Case Is = "Order #3"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$25")
        Case Is = "Order #4"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$36")
        Case Is = "Order #5"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$48")
        Case Is = "Order #6"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$59")
        Case Is = "Order #7"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$71")
        Case Is = "Order #8"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$82")
        Case Is = "Order #9"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$94")
        Case Is = "Order #10"
        Set rng_Orders_Sheet_Home_Cell = ThisWorkbook.Worksheets("Orders").Range("$A$105")
End Select
    str_Opt_Order = rng_Orders_Sheet_Home_Cell.value
    str_Customer_Order_Home_Cell = rng_Orders_Sheet_Home_Cell.Address
    var_Customer_Order_Home_Cell = rng_Orders_Sheet_Home_Cell.Address
    mod_Order_Form_Maintenance.opt_Order_Buttons_Click

End Sub

Private Sub opt_Order5_Click()
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
End Sub

Private Sub opt_Order6_Click()

   Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
End Sub

Private Sub opt_Order7_Click()
  Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
End Sub

Private Sub opt_Order8_Click()
    
   Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
    
End Sub

Private Sub opt_Order9_Click()
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
    End Sub
Private Sub opt_Order10_Click()
    Call Mod_Order_Sheet.handle_Selected_Options_Order_Block(Me)
End Sub

Private Sub txt_Customer_Name_Change()

End Sub





Public Sub initialize_Form_Binders()
    'RUNNING CODE 7-21-2025
    Dim obj_Dict As Object
    Dim fieldKey As Variant
    Dim ctl As MSForms.Control
    Dim i As Long
    
    
    Set obj_Dict = get_Order_Form_Field_Map()
    
    Debug.Print "?? Initializing Form Binders..."
    Debug.Print "Field map count: " & obj_Dict.Count
    
    
    ReDim fieldBinders(0 To obj_Dict.Count - 1)

    i = 0
    For Each fieldKey In obj_Dict.Keys
    
        On Error Resume Next
        Set ctl = Me.Controls(obj_Dict(fieldKey))
        If Err.Number <> 0 Then
            Debug.Print "?? Control '" & obj_Dict(fieldKey) & "' not found in form."
            Err.Clear
            On Error GoTo 0
        Else
    
            Debug.Print "? Found control '" & obj_Dict(fieldKey) & "' ? Type: " & TypeName(ctl)
         If TypeName(ctl) = "TextBox" Or _
                TypeName(ctl) = "ComboBox" Or _
                TypeName(ctl) = "OptionButton" Or _
                TypeName(ctl) = "ListBox" Or _
                TypeName(ctl) = "CheckBox" Then

    Set fieldBinders(i) = New cls_Order_Form_To_Sheet_Binder
    fieldBinders(i).fieldName = fieldKey
    Set fieldBinders(i).boundControl = ctl
    Debug.Print "?? Bound: " & fieldBinders(i).fieldName & " ? " & TypeName(ctl)
 
   
    i = i + 1
Else
    Debug.Print "? Skipped unsupported control type ? " & TypeName(ctl) & " ? " & obj_Dict(fieldKey)
End If

        End If
'        If Not ctl Is Nothing Then
'            Set fieldBinders(i) = New cls_Order_Form_To_Sheet_Binder
'            'Set fieldBinders(i).boundControl = ctl 'Permanent Code
'            Set fieldBinders(i).boundTextBox = ctl  ' Test Code
'            fieldBinders(i).fieldName = fieldKey
'            i = i + 1
'        End If
    Next fieldKey
End Sub



  
  

Private Sub OptionButton_Click()
    Dim ob As OptionButton
    Set ob = Me.ActiveControl
    MsgBox "You clicked on: " & ob.Name
End Sub

Private binderTest As cls_Order_Form_To_Sheet_Binder

Private Sub UserForm_Initialize()

    
    
   Call init_Form_New_Orders(Me)
    Call initialize_Form_Binders

   ' Call init_Form_New_Orders(Me)
 
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'RUNNING CODE 7-21-2025
    If CloseMode = 0 Then  ' User clicked ??
        Debug.Print "?? Form closed via X — triggering safe cleanup"
        IsFormShuttingDown = True
        Call DeactivateBinders
    End If
End Sub
Public Sub DeactivateBinders()
    'RUNNING CODE 7-21-2025
    Dim i As Long
    If IsArray(fieldBinders) Then
        For i = LBound(fieldBinders) To UBound(fieldBinders)
            If Not fieldBinders(i) Is Nothing Then
                With fieldBinders(i)
                    Set .m_boundTextBox = Nothing
                    Set .m_boundComboBox = Nothing
                    Set .m_boundOptionButton = Nothing
                    Set .m_boundListBox = Nothing
                    Set .m_boundCheckBox = Nothing
                End With
                Set fieldBinders(i) = Nothing
            End If
        Next i
    End If
    Erase fieldBinders
    Debug.Print "?? All binders deactivated and erased."
End Sub

