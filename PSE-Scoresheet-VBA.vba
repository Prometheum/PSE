

Sub AddSheets()
'This Macro will copy the copy row into a new sheet, take the next 20 rows and cut them from the original
'and paste them into the new sheet then delete them. It will creat a new file for every 20 rows

Application.EnableEvents = False
Dim wsMasterSheet As Excel.Worksheet
Dim wb As Excel.Workbook
Dim rowCount As Integer
Dim rowsPerSheet As Integer
Dim newBook As Excel.Workbook
Set wsMasterSheet = ActiveSheet
Set wb = ActiveWorkbook

rowsPerSheet = 20
rowCount = Application.CountA(Sheets(1).Range("B:B"))
sheetCount = WorksheetFunction.RoundUp(rowCount / rowsPerSheet, 0)

Dim i As Integer, bundleNum As Integer

For i = 1 To sheetCount Step 1
Set newBook = Workbooks.Add
Set newBookSheet = newBook.Sheets("Sheet1")

bundleNum = bundleNum + 1
With wb
    
    
    
     wsMasterSheet.Range("A1:E1").EntireRow.Copy Destination:=Sheets(.Sheets.Count).Range("A1").End(xlUp)

    wsMasterSheet.Range("A2:" & "C" & (rowsPerSheet + 1)).EntireRow.Cut Destination:=newBook.Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Offset(1)
    wsMasterSheet.Range("A2:" & "C" & (rowsPerSheet + 1)).EntireRow.Delete
      With newBook
        With newBookSheet
        .Name = "PJ-Bundle-" & bundleNum
        .UsedRange.Columns.AutoFit
        .Range("C:C").Locked = False
        .Protect "yes"
        End With
       
         .SaveAs "C:\users\adam1\Documents\PSE\TE-Bundle" & bundleNum & ".xlsx" 'New File name here, Change Division where required
        .Close
        End With
       
        
        
     Set newBook = Nothing
End With


Next

wsMasterSheet.Name = "Rows 1 - " & rowsPerSheet

Application.EnableEvents = True

End Sub