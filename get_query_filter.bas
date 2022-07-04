Option Explicit

Sub query_filter()

Dim lngRow, lngRowMax As Long
Dim lngCounter, lngVarCounter As Long
Dim lngFilter, lngValue As Long
Dim lngCounterValue, lngCounterRange As Long
Dim lngFilterRows, lngRowValue As Long
Dim strPartNumber As String
Dim varFilterRows, varRange, varValue As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1 ' Replace worksheet name

' strPartNumber data is located on column "A" in this example.

lngRowMax = wksSheet.Cells(wksSheet.Rows.Count, 1).End(xlUp).row
lngValue = 0

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
  
lngCounterValue = 0
lngFilterRows = 0
ReDim varFilterRows(lngFilterRows)

For lngCounterValue = LBound(varValue) To UBound(varValue)
    strPartNumber = varValue(lngCounterValue)

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
    
    ' Results of the filter
    lngFilter = wksSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 2).row
    
    ' Get range of unique part numbers
    varRange = get_visible_rows

    ' Go through the list of identical numbers
    For lngCounterRange = LBound(varRange) To UBound(varRange)
      lngRowValue = varRange(lngCounterRange)
      Debug.Print "Row: " & lngRowValue
    Next lngCounterRange

Next lngCounterValue

wksSheet.AutoFilterMode = False

End Sub

' =============================================
Function get_visible_rows() As Variant

Dim lngRow, lngRowMax As Long
Dim index As Long
Dim varRange As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1 ' Replace worksheet name

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
