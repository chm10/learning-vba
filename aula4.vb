Public Sub ImportTextFile()
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    
    OpenFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
    For i = 1 To Application.CountA(OpenFiles)
        Set TextFile = Workbooks.Open(OpenFiles(i))
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        TextFile.Close
    Next i
    
End Sub