Sub reset_Autofilter()

Dim lngColumn, lngColumnMax As Long
Dim lngColumnCategory As Long
Dim strCategory As String
Dim wksSheet As Worksheet

Set wksSheet = Sheet1     ' Replace worksheet name if needed.

strCategory = "Category"  ' Replace strCategory name if needed.

lngColumnMax = wksSheet.UsedRange.Columns.Count

For lngColumn = 1 To lngColumnMax     ' Find the column number of strCategory
  If wksSheet.Cells(1, lngColumn).Value = strCategory Then lngColumnCategory = lngColumn
Next lngColumn

With wksSheet.Range("1:1")
  .AutoFilter Field:=lngColumnCategory     ' This line resets the AutoFilter value.
End With

End Sub
