Sub SortTableAndMove()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rngTable As Range
    Dim rngSort As Range
    Dim rngFilter As Range
    Dim strValue As String
    
    ' Change "Sheet1" to the name of the source worksheet
    Set wsSource = ThisWorkbook.Worksheets("Sheet1")
    
    ' Change "A1" to the top-left cell of the table
    Set rngTable = wsSource.Range("A1").CurrentRegion
    
    ' Change "Column1" to the name of the column that contains the values to filter by
    ' Also change "Value1" to the specific value that you want to filter by
    strValue = "Value1"
    
    ' Change "Sheet2" to the name of the target worksheet
    Set wsTarget = ThisWorkbook.Worksheets("Sheet2")
    
    ' Set the range to sort and filter by
    Set rngSort = rngTable.Sort(Key1:=wsSource.Range("Column1"), Order1:=xlAscending, Header:=xlYes)
    Set rngFilter = rngSort.Columns(1).Cells.SpecialCells(xlCellTypeVisible)
    
    ' Clear the target worksheet
    wsTarget.Cells.Clear
    
    ' Copy the filtered table to the target worksheet
    rngFilter.Copy Destination:=wsTarget.Range("A1")
    
    ' Remove the filters
    wsSource.AutoFilterMode = False
End Sub