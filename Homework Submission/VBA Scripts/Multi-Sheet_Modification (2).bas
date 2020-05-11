Attribute VB_Name = "Module3"
'Allow Script to Run on Each Worksheet

Sub Run_All_Sheets()

'Define Variable as type worksheet
Dim ws As Worksheet

Application.ScreenUpdating = False

For Each ws In Worksheets
    ws.Select
    'Call in the Above Script
    Call VBA_Homework
Next

Application.ScreenUpdating = True

End Sub

