VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Order_Form_To_Sheet_Binder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public WithEvents m_boundTextBox As MSForms.TextBox
Attribute m_boundTextBox.VB_VarHelpID = -1
Public WithEvents m_boundComboBox As MSForms.ComboBox
Attribute m_boundComboBox.VB_VarHelpID = -1
Public WithEvents m_boundOptionButton As MSForms.OptionButton
Attribute m_boundOptionButton.VB_VarHelpID = -1
Public WithEvents m_boundListBox As MSForms.ListBox
Attribute m_boundListBox.VB_VarHelpID = -1
Public WithEvents m_boundCheckBox As MSForms.CheckBox
Attribute m_boundCheckBox.VB_VarHelpID = -1



' ?? Metadata
Public fieldName As String
Public ParentForm As Object
Public wsTarget As Worksheet

' ?? Assignment entry point: route control to correct WithEvents binder
Public Property Set boundControl(ctrl As MSForms.Control)
    If ctrl Is Nothing Then
        Debug.Print "?? Binder rejected null control for field: " & fieldName
        Exit Property
    End If

    Select Case TypeName(ctrl)
        Case "TextBox"
            Set m_boundTextBox = ctrl
            Debug.Print "? TextBox bound ? Field: " & fieldName & " ? Control: " & ctrl.Name

        Case "ComboBox"
            Set m_boundComboBox = ctrl
            Debug.Print "? ComboBox bound ? Field: " & fieldName & " ? Control: " & ctrl.Name

        Case "OptionButton"
            Set m_boundOptionButton = ctrl
            Debug.Print "? OptionButton bound ? Field: " & fieldName & " ? Control: " & ctrl.Name

        Case "ListBox"
            Set m_boundListBox = ctrl
            Debug.Print "? ListBox bound ? Field: " & fieldName & " ? Control: " & ctrl.Name

        Case "CheckBox"
            Set m_boundCheckBox = ctrl
            Debug.Print "? CheckBox bound ? Field: " & fieldName & " ? Control: " & ctrl.Name

        Case Else
            Debug.Print "?? Unsupported control type ? " & TypeName(ctrl) & " ? Field: " & fieldName
    End Select
End Property



' ?? Debug support: report which control type is currently bound
Public Property Get BoundControlType() As String
    If Not m_boundTextBox Is Nothing Then BoundControlType = "TextBox"
    If Not m_boundComboBox Is Nothing Then BoundControlType = "ComboBox"
    If Not m_boundOptionButton Is Nothing Then BoundControlType = "OptionButton"
    If Not m_boundListBox Is Nothing Then BoundControlType = "ListBox"
    If Not m_boundCheckBox Is Nothing Then BoundControlType = "CheckBox"
    If BoundControlType = "" Then BoundControlType = "(none)"
End Property


 'Print TextBox; Change; handler
Private Sub m_boundTextBox_Change()
    If frm_New_Orders.boo_Is_Bulk_Clearing Or IsFormShuttingDown Then Exit Sub

    Call liveWrite_FieldValue(fieldName, m_boundTextBox.Text)
End Sub

' ?? ComboBox change handler
Private Sub m_boundComboBox_Change()
    If frm_New_Orders.boo_Is_Bulk_Clearing Or IsFormShuttingDown Then Exit Sub
    Debug.Print "?? ComboBox '" & fieldName & "' changed ? " & m_boundComboBox.Text
    Call liveWrite_FieldValue(fieldName, m_boundComboBox.Text)
End Sub

' ? OptionButton click handler
Private Sub m_boundOptionButton_Click()
   If frm_New_Orders.boo_Is_Bulk_Clearing Or IsFormShuttingDown Then Exit Sub
    Debug.Print "? OptionButton '" & fieldName & "' clicked ? " & m_boundOptionButton.value
    Call liveWrite_FieldValue(fieldName, m_boundOptionButton.value)
End Sub
Private Sub m_boundCheckBox_Click()
    If frm_New_Orders.boo_Is_Bulk_Clearing Or IsFormShuttingDown Then Exit Sub
    Debug.Print "?? CheckBox '" & fieldName & "' clicked ? " & m_boundCheckBox.value
    Call liveWrite_FieldValue(fieldName, m_boundCheckBox.value)
End Sub
Private Sub m_boundListBox_Change()
    If frm_New_Orders.boo_Is_Bulk_Clearing Or IsFormShuttingDown Then Exit Sub
    Debug.Print "?? ListBox '" & fieldName & "' changed ? " & m_boundListBox.value
    Call liveWrite_FieldValue(fieldName, m_boundListBox.value)
End Sub

Private Sub m_boundListBox_Click()
    If frm_New_Orders.boo_Is_Bulk_Clearing Or IsFormShuttingDown Then Exit Sub
    Debug.Print "?? ListBox '" & fieldName & "' changed ? " & m_boundListBox.value
    Call liveWrite_FieldValue(fieldName, m_boundListBox.value)
End Sub


'
'Private WithEvents m_boundTextBox As MSForms.TextBox
'Private WithEvents m_boundComboBox As MSForms.ComboBox
'Private WithEvents m_boundOptionButton As MSForms.OptionButton
'
'
'
'Public fieldName As String
'Public ParentForm As Object
'Public wsTarget As Worksheet
'Public Property Set boundControl(ctrl As MSForms.Control)
'    If ctrl Is Nothing Then Exit Property
'
'    Select Case TypeName(ctrl)
'        Case "TextBox": Set m_boundTextBox = ctrl
'        Case "ComboBox": Set m_boundComboBox = ctrl
'        Case "OptionButton": Set m_boundOptionButton = ctrl
'        Case Else: Debug.Print "?? Unsupported control type: " & TypeName(ctrl)
'    End Select
'End Property
'Private Sub m_boundTextBox_Change()
'    Debug.Print "?? TextBox '" & fieldName & "' changed ? " & m_boundTextBox.Text
'    Call liveWrite_FieldValue(fieldName, m_boundTextBox.Text)
'End Sub
'
'Private Sub m_boundComboBox_Change()
'    Debug.Print "?? ComboBox '" & fieldName & "' changed ? " & m_boundComboBox.Text
'    Call liveWrite_FieldValue(fieldName, m_boundComboBox.Text)
'End Sub
'
'Private Sub m_boundOptionButton_Click()
'    Debug.Print "? OptionButton '" & fieldName & "' clicked ? " & m_boundOptionButton.value
'    Call liveWrite_FieldValue(fieldName, m_boundOptionButton.value)
'End Sub



''Private Sub boundControl_Change()
'Private Sub m_boundControl_Change()
'    Debug.Print "? Change fired ? " & fieldName & " = " & boundControl.Text
'
'    If TypeName(boundControl) = "TextBox" Or TypeName(boundControl) = "ComboBox" Or TypeName(boundControl) = "ListBox" Then
'        Call liveWrite_FieldValue(fieldName, boundControl.Text)
'    End If
'End Sub
'
'
''Private Sub boundControl_Click()
'Private Sub m_boundControl_Click()
'    ' For option buttons or clickable inputs
'    If TypeName(boundControl) Like "OptionButton*" Then
'        Call liveWrite_FieldValue(fieldName, boundControl.value)
'    End If
'End Sub
'
'Private Sub boundTextBox_Change()
'    Debug.Print "? TextBox change triggered for ? " & fieldName
'End Sub
'
'
'Public Property Set boundControl(ctrl As MSForms.Control)
'    If TypeName(ctrl) = "TextBox" Then
'        Set m_boundTextBox = ctrl
'    Else
'        Debug.Print "?? Unsupported control type: " & TypeName(ctrl)
'    End If
'End Property
'
''
''
'Private Sub m_boundTextBox_Change()
'    Debug.Print "?? Change fired ? " & fieldName & " = " & m_boundTextBox.Text
'    Call liveWrite_FieldValue(fieldName, m_boundTextBox.Text)
'End Sub
''
'
'Private Sub m_boundComboBox_Change()
'    Debug.Print "?? ComboBox '" & fieldName & "' changed ? " & m_boundComboBox.Text
'End Sub
'
'Private Sub m_boundOptionButton_Click()
'    Debug.Print "? OptionButton '" & fieldName & "' clicked ? " & m_boundOptionButton.value
'End Sub


