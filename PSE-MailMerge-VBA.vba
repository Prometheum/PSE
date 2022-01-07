Option Explicit

Const FOLDER_SAVED As String = "C:\Users\adam1\Documents\PSE\IS\" 'Makes sure your folder path ends with a backward slash
Const SOURCE_FILE_PATH As String = "C:\Users\adam1\Downloads\IS.xls" 'File used in the mailmerge, provides PSE data

Sub TestRun()
Dim MainDoc As Document, TargetDoc As Document
Dim dbPath As String, bundle As String
Dim recordNumber As Long, totalRecord As Long
Dim startNum As Long, endNum As Long, bundleNum As Long


Set MainDoc = ActiveDocument
    With MainDoc.MailMerge
    
        '// if you want to specify your data, insert a WHERE clause in the SQL statement
        .OpenDataSource Name:=SOURCE_FILE_PATH, sqlstatement:="SELECT * FROM [Sheet1$]"
            
        totalRecord = .DataSource.RecordCount
        
        If startNum < totalRecord Then
          If endNum > totalRecord Then
             endNum = totalRecord
          End If
          bundleNum = 1
bundle = "IS-Bundle" 'Change this bundle string to the division you are printing, IS/PJ/TECH
endNum = 20 'Cuts the file into 20s
startNum = 1
        
       
        
        For recordNumber = 1 To (totalRecord / 20) + 1
            With .DataSource
                .ActiveRecord = startNum
                .FirstRecord = startNum
                .LastRecord = endNum
            End With
            
                .Destination = wdSendToNewDocument
                .Execute False
               
    
                With ActiveDocument
                .SaveAs FileName:=FOLDER_SAVED & bundle & bundleNum & ".docx", FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=False
                .ExportAsFixedFormat OutputFileName:= _
                   FOLDER_SAVED & bundle & bundleNum & ".pdf", _
                   ExportFormat:=wdExportFormatPDF, _
                   OpenAfterExport:=False, _
                   OptimizeFor:=wdExportOptimizeForPrint, _
                   Range:=wdExportAllDocument, _
                   IncludeDocProps:=True, _
                   CreateBookmarks:=wdExportCreateWordBookmarks, _
                    BitmapMissingFonts:=True
                .Close
                
                
                End With
                
               
                endNum = endNum + 20
                startNum = startNum + 20
                bundleNum = bundleNum + 1
                If startNum > totalRecord Then
                Stop
                End If
                
                Next recordNumber
                
        End If
    End With
    
 Set MainDoc = Nothing

End Sub










