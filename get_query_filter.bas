Option Explicit

Sub query_filter()

Dim lngFilter As Long
Dim lngCounterRange, lngCounterValue As Long
Dim lngRowValue As Long
Dim strPartNumber As String
Dim varRange, varValue As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

' Create an array of unique part numbers:
varValue = create_array
        
' Go through the unique part numbers
For lngCounterValue = LBound(varValue) To UBound(varValue)
    strPartNumber = varValue(lngCounterValue)
    
    ' Apply the filter
    Call apply_filter(strPartNumber)
    
    ' Results of the filter
    lngFilter = wksSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).row
    
    ' Get range of unique part numbers
    varRange = get_visible_rows

    ' Go through the list of identical numbers
    For lngCounterRange = LBound(varRange) To UBound(varRange)
      lngRowValue = varRange(lngCounterRange)
      Debug.Print "Row: " & lngRowValue & " " & strPartNumber
    Next lngCounterRange

Next lngCounterValue

wksSheet.AutoFilterMode = False

End Sub

' ================================================================
Sub apply_filter(ByVal strPartNumber As String)

Dim wksSheet As Worksheet
Set wksSheet = Sheet1

With wksSheet.Rows("1:1")
    .AutoFilter Field:=1, Criteria1:=strPartNumber
End With

wksSheet.AutoFilter.Sort.SortFields. _
    Add2 Key:=Range("A2"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal

With wksSheet.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Sub

' ================================================================
Function create_array() As Variant

Dim lngRow, lngRowMax, lngValue As Long
Dim strPartNumber As String
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

lngRowMax = wksSheet.Cells(wksSheet.Rows.Count, 1).End(xlUp).row

ReDim varValue(lngValue)

For lngRow = 2 To lngRowMax
' Read available values, e.g. part numbers without duplicates
    strPartNumber = wksSheet.Cells(lngRow, 1).Value
    If Not IsNumeric(Application.Match(strPartNumber, varValue, 0)) Then ' No duplicates
      varValue(lngValue) = strPartNumber
      lngValue = lngValue + 1
      ReDim Preserve varValue(lngValue)
    End If
Next lngRow

' Delete the last created empty slot of the current array.
If varValue(UBound(varValue)) = "" Then
  ReDim Preserve varValue(UBound(varValue) - 1)
End If

create_array = varValue

End Function

' ================================================================
Function get_visible_rows() As Variant

Dim lngRow, lngRowMax As Long
Dim index As Long
Dim varRange As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

index = 0
lngRowMax = wksSheet.UsedRange.Rows.Count

ReDim varRange(index)
For lngRow = 2 To lngRowMax
  If Not wksSheet.Rows(lngRow).EntireRow.Hidden Then
    varRange(index) = lngRow
    index = index + 1
    ReDim Preserve varRange(index)
  End If
Next

ReDim Preserve varRange(index - 1)
get_visible_rows = varRange

End Function
