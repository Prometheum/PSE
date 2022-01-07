Sub formatSheets()
'
' add page breaks every 20 rows until it gets to the 293 row

For i = 20 To 548 Step 20
    ActiveSheet.HPageBreaks.Add Before:=Cells(i + 1, 1)
Next
End Sub