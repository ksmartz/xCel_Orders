Attribute VB_Name = "mod_Maybe_Use"
Public Sub GetSeriesMeta(ByVal selectedSeries As String)
    Dim ws As Worksheet, r As Long, valA As String, valB As String, valC As String

    Set ws = ThisWorkbook.Sheets(str_Manufacturer_Name)

    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        valA = Trim(ws.Cells(r, 1).value)
        If valA = selectedSeries Then
            valB = Trim(ws.Cells(r, 2).value)
            valC = Trim(ws.Cells(r, 3).value)

            Debug.Print "Subcategory: " & valB
            Debug.Print "Notes: " & valC
            Exit Sub
        End If
    Next r
End Sub

