Option Explicit

Sub query_filter()

Dim lngCounterValue As Long
Dim strPartnumber As String
Dim varRange, varValue As Variant
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

If Sheet1.Cells(1, 1).Value = "" Then Sheet1.Cells(1, 1).Value = "Data"

' Create an array of unique part numbers:
varValue = create_array(wksSheet)
        
' Go through the unique part numbers
For lngCounterValue = LBound(varValue) To UBound(varValue)
    strPartnumber = varValue(lngCounterValue)
    
    ' Apply the filter
    Call apply_filter(strPartnumber, wksSheet)
        
    ' Get range of unique part numbers
    varRange = get_visible_rows(wksSheet)

    ' Go through the list of numbers to be checked
    Call check_list(strPartnumber, varRange)

Next lngCounterValue

wksSheet.AutoFilterMode = False

End Sub

' ================================================================
Function create_array(ByRef wksSheet As Worksheet) As Variant

Dim lngRow, lngRowMax, lngValue As Long
Dim strPartnumber As String

lngRowMax = wksSheet.Cells(wksSheet.Rows.Count, 1).End(xlUp).Row

ReDim varValue(lngValue)

For lngRow = 2 To lngRowMax
' Read available values, e.g. part numbers without duplicates
    strPartnumber = wksSheet.Cells(lngRow, 1).Value
    If Not IsNumeric(Application.Match(strPartnumber, varValue, 0)) Then ' No duplicates
      varValue(lngValue) = strPartnumber
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
Sub apply_filter(ByVal strPartnumber As String, ByRef wksSheet As Worksheet)

With wksSheet.Rows("1:1")
    .AutoFilter Field:=1, Criteria1:=strPartnumber
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
Function get_visible_rows(ByRef wksSheet As Worksheet) As Variant

Dim lngIndex, lngRow, lngRowMax As Long
Dim varRange As Variant

lngRowMax = wksSheet.UsedRange.Rows.Count

ReDim varRange(lngIndex)
For lngRow = 2 To lngRowMax
  If Not wksSheet.Rows(lngRow).EntireRow.Hidden Then
    varRange(lngIndex) = lngRow
    lngIndex = lngIndex + 1
    ReDim Preserve varRange(lngIndex)
  End If
Next

ReDim Preserve varRange(lngIndex - 1)

get_visible_rows = varRange

End Function

' ================================================================
Sub check_list(ByVal strPartnumber As String, ByRef varRange As Variant)

Dim lngCounterRange, lngRowValue As Long

For lngCounterRange = LBound(varRange) To UBound(varRange)
  lngRowValue = varRange(lngCounterRange)
  Debug.Print "Row: " & lngRowValue & " " & strPartnumber
Next lngCounterRange

End Sub
