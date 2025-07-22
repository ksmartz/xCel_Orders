Attribute VB_Name = "Mod_Git_Hub"

Sub ExportAllVBAModules()
    Dim vbComp As Object
    Dim exportPath As String
    ' Set the export folder path
    exportPath = "C:\Users\ksmar\OneDrive\Documents\GH-XCel_Orders_Workbook\Exports\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    ' Loop through all components in the VBA project
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 2 ' Class Module
                vbComp.Export exportPath & vbComp.Name & ".cls"
            Case 3 ' UserForm
                vbComp.Export exportPath & vbComp.Name & ".frm"
            Case Else
                ' Skip document modules like ThisWorkbook or Sheet1
        End Select
    Next vbComp

    MsgBox "Export complete! Files saved to: " & exportPath
End Sub

