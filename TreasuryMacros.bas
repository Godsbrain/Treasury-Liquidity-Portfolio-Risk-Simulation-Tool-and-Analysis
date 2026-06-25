Attribute VB_Name = "Module1"
Sub RefreshDashboard()
    ' Prevents screen flickering while the code runs
    Application.ScreenUpdating = False
    
    ' Calculates all formulas in the workbook
    Calculate
    
    ' Refreshes Pivot Tables on all sheets
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Treasury Forecast successfully updated and data refreshed.", vbInformation, "System Update"
End Sub

Sub ApplyScenario(ScenarioType As String)
    ' UPDATED: Reference to "Scenario Summary"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Scenario Summary")
    
    Select Case ScenarioType
        Case "Positive"
            ws.Range("B2").Value = 400
            ws.Range("B4").Value = 450
        Case "Realistic"
            ws.Range("B2").Value = 345
            ws.Range("B4").Value = 400
        Case "Pessimistic"
            ws.Range("B2").Value = 300
            ws.Range("B4").Value = 350
    End Select
    
    Calculate
    MsgBox ScenarioType & " scenario applied successfully.", vbInformation, "Scenario Applied"
End Sub

' Helper macros
Sub Scenario_Positive(): ApplyScenario "Positive": End Sub
Sub Scenario_Realistic(): ApplyScenario "Realistic": End Sub
Sub Scenario_Pessimistic(): ApplyScenario "Pessimistic": End Sub

Sub ExportDashboardToPDF()
    Dim ws As Worksheet
    ' UPDATED: Reference to "Report" (or "Dashboard" if that's where your charts live)
    ' If your dashboard is on the "Report" sheet, use "Report" below:
    Set ws = ThisWorkbook.Sheets("Report")
    
    ws.PageSetup.PrintArea = "A1:D30"
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & "\Treasury_Report_" & Format(Date, "yyyy-mm-dd") & ".pdf", _
        Quality:=xlQualityStandard
        
    MsgBox "Report exported to PDF successfully.", vbInformation, "Export Complete"
End Sub

