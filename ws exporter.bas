Attribute VB_Name = "Module2"
Sub Save_each_worksheet()


Dim i As Integer
Dim ws As Worksheet
Dim wb As Workbook
Dim savePath As String
Dim filename As String

savePath = ThisWorkbook.Path & "\"

Set wb = ThisWorkbook

For Each ws In wb.Worksheets
  
    If ws.Name <> "Master Data" Then
      
        filename = Left(wb.Name, Len(wb.Name) - 5) & " - " & ws.Index - 1
        
        ws.Copy
        ActiveWorkbook.SaveAs filename:=savePath & filename, FileFormat:=wb.FileFormat
        ActiveWorkbook.Close SaveChanges:=False
    End If
Next ws

End Sub



