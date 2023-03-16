Attribute VB_Name = "Module3"
Sub Delete_ws_records()

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False 'disable alerts
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Master Data" Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True 're-enable alerts

End Sub
