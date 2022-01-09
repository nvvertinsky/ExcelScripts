Sub apostrof()
  Dim cell As Range
  Application.ScreenUpdating = False
  
  For Each cell In Selection.SpecialCells(xlCellTypeConstants)
    cell.Value = " '" & cell.Value & "'"
  Next
  
  Application.ScreenUpdating = True
End Sub  