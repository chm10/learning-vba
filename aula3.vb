Public Sub ImportTextFile()
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    OpenFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
    Set TextFile = Workbooks.Open(OpenFiles(1))
    TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
    Workbooks(1).Activate
    ActiveSheet.Paste
    TextFile.Close
    
End Sub
