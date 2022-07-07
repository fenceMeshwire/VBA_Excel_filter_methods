Option Explicit

Sub filter_Category_1()

Call check_AutoFilter
Call filter_Category(1)   ' The value to be filtered is the number:int 1 in this example.

End Sub

' =============================================
' Check if AutoFilter is activated
Sub check_AutoFilter()

Dim lngColumnMax As Long
Dim wksSheet As Worksheet

Set wksSheet = Sheet1   ' Set the worksheet to be filtered

lngColumnMax = wksSheet.UsedRange.Columns.Count

With wksSheet
  If .AutoFilterMode Then Exit Sub
  .Range("1:1").AutoFilter    ' Set the range to be filtered. 1:1 is the first row of the worksheet
End With

End Sub

' =============================================
' Category filtering procedure
Sub filter_Category(intCategory As Integer)

Dim lngColumn, lngColumnMax As Long
Dim lngColumnCategory As Long
Dim strCategory As String
Dim wksSheet As Worksheet

Set wksSheet = Sheet1       ' Set the worksheet to be filtered

strCategory = "Category"    ' Set the column name by which to filter

lngColumnMax = wksSheet.UsedRange.Columns.Count

For lngColumn = 1 To lngColumnMax
  If wksSheet.Cells(1, lngColumn).Value = strCategory Then lngColumnCategory = lngColumn
Next lngColumn

' The AutoFilter is set on the first row in this example:
With wksSheet.Range("1:1")
  .AutoFilter Field:=lngColumnCategory, Criteria1:=intCategory
End With

End Sub
