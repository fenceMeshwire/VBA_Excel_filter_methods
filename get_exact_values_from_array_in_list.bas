Option Explicit

Sub get_exact_values_from_array_in_list()

Dim lngRow, lngRowMax As Long
Dim strTypeCode As String
Dim varDat As Variant

' Define elements of the array:
varDat = Array("AB12", "CD13", "34LZ", "U98E")
lngRowMax = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row

' Check type codes line by line:
For lngRow = lngRowMax To 2 Step -1
  strTypeCode = Sheet1.Cells(lngRow, 1).Value
  
  ' Check the type code in the list against the array (see function below)
  If is_type_code(strTypeCode, varDat) = False Then
    Sheet1.Rows(lngRow).Delete
  End If
  
Next lngRow

End Sub

' ================================================================
Function is_type_code(ByVal strTypeCode As String, ByRef varDat) As Boolean

Dim intCounter As Integer

For intCounter = LBound(varDat) To UBound(varDat)

  If varDat(intCounter) = strTypeCode Then
    is_type_code = True
    Exit Function
  End If
  
Next intCounter

is_type_code = False

End Function
