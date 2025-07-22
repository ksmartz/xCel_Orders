VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Order_Review 
   Caption         =   "UserForm1"
   ClientHeight    =   7310
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9590.001
   OleObjectBlob   =   "frm_Order_Review.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Order_Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frmTarget As Object
Public targetCell As Range



Private Sub btn_Delete_Order_Data_Click()
    Dim targetCell As Range
    Set targetCell = Me.targetCell

    If targetCell Is Nothing Then
        MsgBox "? Target cell not set. Deletion aborted.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = targetCell.Worksheet
    Dim r As Long: r = targetCell.MergeArea.Row

    Debug.Print "?? Clearing order block anchored at row: " & r

    ' ?? Clear only if box is unchecked
    If Not chk_Name.value Then ws.Cells(r, 4).ClearContents          ' D1
    If Not chk_Platform_Name.value Then ws.Cells(r, 5).ClearContents ' E1
    If Not chk_Equipment_Type.value Then ws.Cells(r + 1, 7).ClearContents     ' G2
    If Not chk_Manufacturer_Name.value Then ws.Cells(r + 1, 4).ClearContents  ' D2
    If Not chk_Series_Name.value Then ws.Cells(r + 1, 5).ClearContents        ' E2
    If Not chk_Model_Name.value Then ws.Cells(r + 1, 6).ClearContents         ' F2
    If Not chk_Fabric_Type_Name.value Then ws.Cells(r + 2, 3).ClearContents   ' C3
    If Not chk_Color_Name.value Then ws.Cells(r + 2, 5).ClearContents         ' E3

    Unload Me
End Sub



Public Sub PopulateOrderListFromMergedCells()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Orders")
    Dim cell As Range
    Dim orderID As Variant, rAnchor As Long

   Me.lst_Order_Number.Clear
Me.lst_Order_Number.AddItem CStr(rAnchor)


    For Each cell In ws.Range("A1:A200") ' Expand as needed
        If cell.MergeCells Then
            rAnchor = cell.MergeArea.Row
            orderID = Trim(cell.MergeArea.Cells(1, 1).value)

            If IsNumeric(orderID) And CLng(orderID) > 0 Then
                Me.lst_Order_Number.AddItem CStr(rAnchor)
                Debug.Print "? Order found: ID [" & orderID & "] at row " & rAnchor
            End If
        End If
    Next cell
End Sub















Private Sub btn_Keep_Order_Data_Click()
    Dim frmTarget As Object: Set frmTarget = Me.frmTarget
    If frmTarget Is Nothing Then
        MsgBox "? Target form reference not found.", vbExclamation
        Exit Sub
    End If

    Dim syncMap As Variant: syncMap = get_Review_Sync_Map()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Orders")
    Dim r As Long: r = GetOrderStartRow(GetSelectedOrderIndex(frmTarget))
    If r < 1 Then Exit Sub

    Dim i As Long, ctrlName As String, chkName As String, lblName As String
    Dim colIndex As Long, keepValue As Boolean
    Dim ctrl As Object, sourceValue As String

    Dim rowOffset As Long, targetRow As Long

For i = LBound(syncMap) To UBound(syncMap)
    ctrlName = syncMap(i)(0)
    chkName = syncMap(i)(1)
    lblName = syncMap(i)(2)
    rowOffset = syncMap(i)(3)
    colIndex = syncMap(i)(4)
    targetRow = r + rowOffset


        keepValue = Me.Controls(chkName).value
        sourceValue = Me.Controls(lblName).Caption

        Set ctrl = Nothing
        On Error Resume Next
        Set ctrl = frmTarget.Controls(ctrlName)
        On Error GoTo 0

        If keepValue Then
            If Not ctrl Is Nothing Then ApplyField ctrl, sourceValue, True
            ws.Cells(targetRow, colIndex).value = sourceValue
        Else
            ws.Cells(targetRow, colIndex).value = ""
        End If
    Next i

    Sync_Review_Order_From_To_Order_Sheet_Literals_Only ws, r
    Unload Me
End Sub
Public Sub PopulatePreviewData()
    Dim rRow As Long
    rRow = CLng(Me.lst_Order_Number.value)

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Orders")

    ' ?? Fill label fields — update control names as needed
    Me.lbl_Name.Caption = ws.Cells(rRow, 3).value  ' Column C = customer name
'    Me.lbl_Product.Caption = ws.Cells(rRow, 4).value   ' Column D = product
'    Me.lbl_Quantity.Caption = ws.Cells(rRow, 5).value  ' Column E = quantity
'    Me.lbl_Status.Caption = ws.Cells(rRow, 6).value    ' Column F = order status

    Debug.Print "? Populated preview for row " & rRow
End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    Call PopulateOrderListFromMergedCells
End Sub

Public Sub SetOrderRow(ByVal r As Long)
    ' ?? Diagnostic output — see what's in the list
    Debug.Print "?? Trying to match worksheet row: " & r
    Dim i As Long
    For i = 0 To Me.lst_Order_Number.ListCount - 1
        Debug.Print "?? List item " & i & ": " & Me.lst_Order_Number.List(i)
    Next i

    ' ?? Match row to list index
    Dim index As Long
    index = OrderIndexFromRow(r)

    If index > -1 Then
        Me.lst_Order_Number.ListIndex = index
        Call PopulatePreviewData  ' Or your preferred loader method
    Else
        Debug.Print "? No matching index found for row " & r
    End If
End Sub


Public Function OrderIndexFromRow(ByVal r As Long) As Long
    Dim i As Long
    For i = 0 To Me.lst_Order_Number.ListCount - 1
        If Me.lst_Order_Number.List(i) = CStr(r) Then

            OrderIndexFromRow = i
            Exit Function
        End If
    Next i
    OrderIndexFromRow = -1
End Function

