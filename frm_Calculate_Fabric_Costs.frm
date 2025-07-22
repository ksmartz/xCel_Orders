VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Calculate_Fabric_Costs 
   Caption         =   "Calculate Fabric Costs"
   ClientHeight    =   5880
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5810
   OleObjectBlob   =   "frm_Calculate_Fabric_Costs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Calculate_Fabric_Costs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Calculate_Click()
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


End Sub

Private Sub btn_Close_Click()
    Me.Hide
    
End Sub

