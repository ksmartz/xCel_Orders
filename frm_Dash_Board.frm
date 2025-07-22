VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Dash_Board 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frm_Dash_Board.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Dash_Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Click()
    
End Sub

Private Sub btn_Launch_Frm_New_Orders_Click()
 
    
    Call Mod_Order_Sheet.handle_Btn_New_Orders
End Sub
