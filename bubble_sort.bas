Sub bubble_sort()

Const lower_boundary = 0
Const upper_boundary = 9

Dim bleSwapped As Boolean
Dim intArrayElement(lower_boundary To upper_boundary) As Integer
Dim intCounter, intStorage As Integer
Dim lngCellFree As Long

Randomize ' Use random numbers for filling the array with data.

For intCounter = lower_boundary To upper_boundary
  intArrayElement(intCounter) = Rnd * 100
Next intCounter

' Sorting the array in ascending order
Do

  bleSwapped = False
  
  For intCounter = lower_boundary To upper_boundary - 1
    
    If intArrayElement(intCounter) > intArrayElement(intCounter + 1) Then
      ' Store the result of the comparison:
      intStorage = intArrayElement(intCounter)
      ' Swap the values of the compared pair:
      intArrayElement(intCounter) = intArrayElement(intCounter + 1)
      ' Store the greater value in the following array element:
      intArrayElement(intCounter + 1) = intStorage
      ' Store True if the numbers have been swapped:
      bleSwapped = True
    End If
    
  Next intCounter
  
Loop While bleSwapped

' Write result to table (Sheet1)
Sheet1.UsedRange.Clear
lngCellFree = 1

For intCounter = lower_boundary To upper_boundary
  Sheet1.Cells(lngCellFree, 1).Value = intArrayElement(intCounter)
  lngCellFree = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row + 1
Next intCounter

End Sub
