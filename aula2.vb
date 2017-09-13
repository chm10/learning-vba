
Public Sub ImportTextFile()
    Dim TextFile As Workbook
    Set TextFile = Workbooks.Open("C:\Users\MAEKAWA\Documents\projeto 7\demo\April2015Sales.txt")
    TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
    Workbooks(1).Activate
    ActiveSheet.Paste
    TextFile.Close
    
End Sub
