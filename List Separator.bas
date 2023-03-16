Attribute VB_Name = "Module1"
Sub List_Separator()

Dim inputRange As Range
Dim numRows As Long
Dim headerRange As Range
Dim numSheets As Integer
Dim sheetNum As Integer
Dim i As Long
Dim j As Long

On Error Resume Next
Set inputRange = Application.InputBox("Enter the range of the list to be separated", Type:=8)
On Error GoTo 0

numRows = InputBox("Enter the number of rows per worksheet")

Set headerRange = inputRange.Rows(1)

numSheets = Int((inputRange.Rows.Count - 1) / numRows) + 1


For sheetNum = 1 To numSheets
    
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Sheet" & sheetNum
    
   
    headerRange.Copy Destination:=Worksheets("Sheet" & sheetNum).Range("A1")
    
  
    For i = 1 To numRows
        j = (sheetNum - 1) * numRows + i
        If j > inputRange.Rows.Count Then Exit For
        inputRange.Rows(j).Copy Destination:=Worksheets("Sheet" & sheetNum).Range("A" & i + 1)
    Next i
Next sheetNum

End Sub



