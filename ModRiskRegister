'================================================================================
' AERPA v10.2 - COMPLETE RISK REGISTER MODULE (14-COLUMN SCHEMA - FINAL)
' Production Grade | Enterprise Ready
' Date: January 17, 2026
'================================================================================

Option Explicit

'================================================================================
' HELPER FUNCTIONS: Min & Max
'================================================================================

Private Function Min(value1 As Double, value2 As Double) As Double
    If value1 < value2 Then
        Min = value1
    Else
        Min = value2
    End If
End Function

Private Function Max(value1 As Double, value2 As Double) As Double
    If value1 > value2 Then
        Max = value1
    Else
        Max = value2
    End If
End Function

'================================================================================
' TABLE FORMATTING
'================================================================================

Public Sub FormatRISK_REGISTERTable()
    Dim wsRR As Worksheet, lastRow As Long, lastCol As Long
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    lastCol = 14  ' A through N
    
    Call FormatRISK_REGISTERHeader(wsRR, lastCol)
    
    If lastRow > 1 Then
        Call ApplyRISK_REGISTERConditionalFormatting(wsRR, lastRow, lastCol)
        Call FormatRISK_REGISTERDataRows(wsRR, lastRow, lastCol)
    End If
    
    Call AddAutoFilterToRISK_REGISTER(wsRR, lastRow, lastCol)
    Call AdjustRISK_REGISTERColumnWidths(wsRR)
    
    MsgBox "Risk Register formatted | Rows: " & lastRow - 1, vbInformation, "AERPA v10.2"
    Exit Sub
ErrorHandler:
    MsgBox "Format error: " & Err.Description, vbCritical
End Sub

Private Sub FormatRISK_REGISTERHeader(wsRR As Worksheet, lastCol As Long)
    Dim headerRange As Range
    Set headerRange = wsRR.Range(wsRR.Cells(1, 1), wsRR.Cells(1, lastCol))
    
    With headerRange
        .Interior.color = RGB(41, 84, 115)
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With
    
    ' 14-column final schema (A1:N1)
    wsRR.Cells(1, 1).value = "Timestamp"
    wsRR.Cells(1, 2).value = "Batch_ID"
    wsRR.Cells(1, 3).value = "Tenant_ID"
    wsRR.Cells(1, 4).value = "Risk_Score"
    wsRR.Cells(1, 5).value = "Confidence"
    wsRR.Cells(1, 6).value = "Driver1"
    wsRR.Cells(1, 7).value = "Driver2"
    wsRR.Cells(1, 8).value = "Driver3"
    wsRR.Cells(1, 9).value = "Recommendation"
    wsRR.Cells(1, 10).value = "Status"
    wsRR.Cells(1, 11).value = "Review_Notes"
    wsRR.Cells(1, 12).value = "Reviewed_By"
    wsRR.Cells(1, 13).value = "Equipment_ID"
    wsRR.Cells(1, 14).value = "Supplier_ID_Encoded"
    
    wsRR.Rows(1).Locked = True
End Sub

Private Function GetRiskScoreColor(riskScore As Double) As Long
    If riskScore >= 75 Then
        GetRiskScoreColor = RGB(255, 100, 100)      ' Critical - red
    ElseIf riskScore >= 60 Then
        GetRiskScoreColor = RGB(255, 200, 100)      ' High - orange
    ElseIf riskScore >= 45 Then
        GetRiskScoreColor = RGB(255, 255, 150)      ' Medium - yellow
    Else
        GetRiskScoreColor = RGB(150, 255, 100)      ' Low - green
    End If
End Function

Private Sub FormatRISK_REGISTERDataRows(wsRR As Worksheet, lastRow As Long, lastCol As Long)
    Dim i As Long, riskScore As Double, riskColor As Long, rowRange As Range
    
    For i = 2 To lastRow
        Set rowRange = wsRR.Range(wsRR.Cells(i, 1), wsRR.Cells(i, lastCol))
        
        ' Alternating row colors
        If i Mod 2 = 0 Then
            rowRange.Interior.color = RGB(242, 242, 242)
        Else
            rowRange.Interior.color = RGB(255, 255, 255)
        End If
        
        ' Risk_Score formatting (Column D = 4)
        riskScore = CDbl(wsRR.Cells(i, 4).value)
        riskColor = GetRiskScoreColor(riskScore)
        With wsRR.Cells(i, 4)
            .Interior.color = riskColor
            .Font.Bold = True
            .NumberFormat = "0.0"
        End With
        
        ' Confidence formatting (Column E = 5, 0-1 scale)
        With wsRR.Cells(i, 5)
            .NumberFormat = "0.0%"
            .HorizontalAlignment = xlCenter
        End With
        
        ' Recommendation formatting (Column I = 9)
        With wsRR.Cells(i, 9)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        ' Status formatting (Column J = 10)
        With wsRR.Cells(i, 10)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        ' Add borders
        With rowRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .color = RGB(200, 200, 200)
        End With
    Next i
End Sub

Private Sub ApplyRISK_REGISTERConditionalFormatting(wsRR As Worksheet, lastRow As Long, lastCol As Long)
    Dim riskScoreRange As Range, statusRange As Range
    
    ' Risk_Score conditional formatting (Column D = 4)
    Set riskScoreRange = wsRR.Range(wsRR.Cells(2, 4), wsRR.Cells(lastRow, 4))
    With riskScoreRange.FormatConditions
        .Delete
        .Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=75"
        .Item(1).Interior.color = RGB(255, 100, 100)
        .Item(1).Font.color = RGB(255, 255, 255)
        .Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=60", Formula2:="=74.9"
        .Item(2).Interior.color = RGB(255, 200, 100)
        .Add Type:=xlCellValue, Operator:=xlBetween, Formula1:="=45", Formula2:="=59.9"
        .Item(3).Interior.color = RGB(255, 255, 150)
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=45"
        .Item(4).Interior.color = RGB(150, 255, 100)
    End With
    
    ' Status conditional formatting (Column J = 10)
    Set statusRange = wsRR.Range(wsRR.Cells(2, 10), wsRR.Cells(lastRow, 10))
    With statusRange.FormatConditions
        .Delete
        .Add Type:=xlExpression, Formula1:="=$J2=""HOLD"""
        .Item(1).Interior.color = RGB(255, 100, 100)
        .Item(1).Font.color = RGB(255, 255, 255)
        .Add Type:=xlExpression, Formula1:="=$J2=""REVIEW"""
        .Item(2).Interior.color = RGB(255, 255, 150)
        .Add Type:=xlExpression, Formula1:="=$J2=""PASS"""
        .Item(3).Interior.color = RGB(150, 255, 100)
    End With
End Sub

Private Sub AddAutoFilterToRISK_REGISTER(wsRR As Worksheet, lastRow As Long, lastCol As Long)
    On Error Resume Next
    wsRR.AutoFilterMode = False
    On Error GoTo 0
    wsRR.Range(wsRR.Cells(1, 1), wsRR.Cells(lastRow, lastCol)).AutoFilter
End Sub

Private Sub AdjustRISK_REGISTERColumnWidths(wsRR As Worksheet)
    wsRR.Columns(1).ColumnWidth = 18   ' Timestamp
    wsRR.Columns(2).ColumnWidth = 12   ' Batch_ID
    wsRR.Columns(3).ColumnWidth = 10   ' Tenant_ID
    wsRR.Columns(4).ColumnWidth = 12   ' Risk_Score
    wsRR.Columns(5).ColumnWidth = 12   ' Confidence
    wsRR.Columns(6).ColumnWidth = 18   ' Driver1
    wsRR.Columns(7).ColumnWidth = 18   ' Driver2
    wsRR.Columns(8).ColumnWidth = 18   ' Driver3
    wsRR.Columns(9).ColumnWidth = 24   ' Recommendation
    wsRR.Columns(10).ColumnWidth = 12  ' Status
    wsRR.Columns(11).ColumnWidth = 25  ' Review_Notes
    wsRR.Columns(12).ColumnWidth = 15  ' Reviewed_By
    wsRR.Columns(13).ColumnWidth = 14  ' Equipment_ID
    wsRR.Columns(14).ColumnWidth = 16  ' Supplier_ID_Encoded
    
    wsRR.Range("B2").Select
    ActiveWindow.FreezePanes = True
End Sub

'================================================================================
' SORT & FILTER OPERATIONS
'================================================================================

Public Sub SortRISK_REGISTERByRiskScore()
    Dim wsRR As Worksheet, lastRow As Long, dataRange As Range
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    Set dataRange = wsRR.Range(wsRR.Cells(1, 1), wsRR.Cells(lastRow, 14))
    With wsRR.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsRR.Columns(4), SortOn:=xlSortOnValues, Order:=xlDescending
        .SetRange dataRange
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Sorted by Risk Score (highest first)", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Sort error: " & Err.Description, vbCritical
End Sub

Public Sub FilterRISK_REGISTERByCritical()
    Dim wsRR As Worksheet, lastRow As Long, filterRange As Range
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    If Not wsRR.AutoFilterMode Then
        Set filterRange = wsRR.Range(wsRR.Cells(1, 1), wsRR.Cells(lastRow, 14))
        filterRange.AutoFilter
    End If
    
    wsRR.Range(wsRR.Cells(1, 4), wsRR.Cells(lastRow, 4)).AutoFilter Field:=4, Criteria1:=">=75"
    MsgBox "Filtered: CRITICAL (Score >= 75)", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Filter error: " & Err.Description, vbCritical
End Sub

Public Sub ClearRISK_REGISTERFilter()
    Dim wsRR As Worksheet
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    If wsRR.AutoFilterMode Then
        wsRR.AutoFilterMode = False
        wsRR.Cells.AutoFilter
    End If
    MsgBox "Filter cleared", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'================================================================================
' HIGHLIGHTING & ANALYSIS
'================================================================================

Public Sub HighlightRiskByDriver(driverName As String)
    Dim wsRR As Worksheet, lastRow As Long, i As Long, highlightCount As Long
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    wsRR.Range(wsRR.Cells(2, 1), wsRR.Cells(lastRow, 14)).Interior.ColorIndex = xlNone
    
    For i = 2 To lastRow
        If UCase(wsRR.Cells(i, 6).value) = UCase(driverName) Or _
           UCase(wsRR.Cells(i, 7).value) = UCase(driverName) Or _
           UCase(wsRR.Cells(i, 8).value) = UCase(driverName) Then
            wsRR.Range(wsRR.Cells(i, 1), wsRR.Cells(i, 14)).Interior.color = RGB(255, 255, 200)
            highlightCount = highlightCount + 1
        End If
    Next i
    
    MsgBox "Highlighted " & highlightCount & " rows", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub HighlightOutliers()
    Dim wsRR As Worksheet, lastRow As Long, i As Long
    Dim riskScores() As Double, mean As Double, stdDev As Double, zScore As Double
    Dim outlierCount As Long, n As Long, sum As Double, sumSq As Double
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    n = lastRow - 1
    ReDim riskScores(1 To n)
    
    For i = 2 To lastRow
        riskScores(i - 1) = CDbl(wsRR.Cells(i, 4).value)
    Next i
    
    For i = 1 To n
        sum = sum + riskScores(i)
    Next i
    mean = sum / n
    
    For i = 1 To n
        sumSq = sumSq + (riskScores(i) - mean) ^ 2
    Next i
    stdDev = Sqr(sumSq / (n - 1))
    
    wsRR.Range(wsRR.Cells(2, 1), wsRR.Cells(lastRow, 14)).Interior.ColorIndex = xlNone
    
    For i = 2 To lastRow
        If stdDev > 0 Then
            zScore = Abs(riskScores(i - 1) - mean) / stdDev
            If zScore > 2.5 Then
                wsRR.Range(wsRR.Cells(i, 1), wsRR.Cells(i, 14)).Interior.color = RGB(255, 150, 150)
                outlierCount = outlierCount + 1
            End If
        End If
    Next i
    
    MsgBox "Highlighted " & outlierCount & " outliers (z-score > 2.5)" & vbCrLf & _
           "Mean: " & Format(mean, "0.0") & " | StdDev: " & Format(stdDev, "0.0"), vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub FlagAnomalousRecords()
    Dim wsRR As Worksheet, lastRow As Long, i As Long
    Dim anomalyCount As Long, riskScore As Double, confidence As Double
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    wsRR.Range(wsRR.Cells(2, 1), wsRR.Cells(lastRow, 14)).Interior.ColorIndex = xlNone
    
    For i = 2 To lastRow
        riskScore = CDbl(wsRR.Cells(i, 4).value)
        confidence = CDbl(wsRR.Cells(i, 5).value)
        
        If riskScore >= 60 And confidence < 0.5 Then
            wsRR.Range(wsRR.Cells(i, 1), wsRR.Cells(i, 14)).Interior.color = RGB(255, 100, 100)
            anomalyCount = anomalyCount + 1
        End If
    Next i
    
    MsgBox "Flagged " & anomalyCount & " anomalies (High Risk + Low Confidence)", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'================================================================================
' EXPORT & REPORTING
'================================================================================

Public Sub ExportRISK_REGISTERFiltered()
    Dim wsRR As Worksheet, filePath As String, fileName As String
    Dim fileNum As Integer, lineText As String, i As Long, j As Long
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    
    fileName = "AERPA_RISK_REGISTER_" & Format(Now(), "yyyymmddhhmmss") & ".csv"
    filePath = ThisWorkbook.Path & "\" & fileName
    fileNum = FreeFile
    
    Open filePath For Output As fileNum
    
    ' Write header
    lineText = ""
    For j = 1 To 14
        lineText = lineText & """" & wsRR.Cells(1, j).value & """"
        If j < 14 Then lineText = lineText & ","
    Next j
    Print #fileNum, lineText
    
    ' Write data rows
    On Error Resume Next
    For i = 2 To wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
        If wsRR.Rows(i).Hidden = False Then
            lineText = ""
            For j = 1 To 14
                lineText = lineText & """" & wsRR.Cells(i, j).value & """"
                If j < 14 Then lineText = lineText & ","
            Next j
            Print #fileNum, lineText
        End If
    Next i
    On Error GoTo ErrorHandler
    
    Close fileNum
    MsgBox "Exported: " & filePath, vbInformation
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Close fileNum
    MsgBox "Export error: " & Err.Description, vbCritical
End Sub

Public Sub BulkExportByStatus()
    Dim wsRR As Worksheet, lastRow As Long, i As Long, j As Long
    Dim holdFile As String, reviewFile As String, passFile As String
    Dim holdNum As Integer, reviewNum As Integer, passNum As Integer
    Dim lineText As String, status As String
    Dim holdCount As Long, reviewCount As Long, passCount As Long
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    holdFile = ThisWorkbook.Path & "\AERPA_HOLD_" & Format(Now(), "yyyymmddhhmmss") & ".csv"
    reviewFile = ThisWorkbook.Path & "\AERPA_REVIEW_" & Format(Now(), "yyyymmddhhmmss") & ".csv"
    passFile = ThisWorkbook.Path & "\AERPA_PASS_" & Format(Now(), "yyyymmddhhmmss") & ".csv"
    
    holdNum = FreeFile
    reviewNum = FreeFile + 1
    passNum = FreeFile + 2
    
    Open holdFile For Output As holdNum
    Open reviewFile For Output As reviewNum
    Open passFile For Output As passNum
    
    ' Write headers
    lineText = ""
    For j = 1 To 14
        lineText = lineText & """" & wsRR.Cells(1, j).value & """"
        If j < 14 Then lineText = lineText & ","
    Next j
    Print #holdNum, lineText
    Print #reviewNum, lineText
    Print #passNum, lineText
    
    ' Write data by status
    For i = 2 To lastRow
        status = UCase(Trim(wsRR.Cells(i, 10).value))
        lineText = ""
        For j = 1 To 14
            lineText = lineText & """" & wsRR.Cells(i, j).value & """"
            If j < 14 Then lineText = lineText & ","
        Next j
        
        Select Case status
            Case "HOLD": Print #holdNum, lineText: holdCount = holdCount + 1
            Case "REVIEW": Print #reviewNum, lineText: reviewCount = reviewCount + 1
            Case "PASS": Print #passNum, lineText: passCount = passCount + 1
        End Select
    Next i
    
    Close holdNum
    Close reviewNum
    Close passNum
    
    MsgBox "Exported:" & vbCrLf & "HOLD: " & holdCount & vbCrLf & _
           "REVIEW: " & reviewCount & vbCrLf & "PASS: " & passCount, vbInformation
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Close holdNum: Close reviewNum: Close passNum
    MsgBox "Export error: " & Err.Description, vbCritical
End Sub

Public Sub GenerateRISK_REGISTERSummary()
    Dim wsRR As Worksheet, lastRow As Long, i As Long
    Dim totalBatches As Long, holdCount As Long, reviewCount As Long, passCount As Long
    Dim avgRiskScore As Double, criticalCount As Long, riskSum As Double
    Dim summaryMsg As String, riskScore As Double, status As String
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    totalBatches = lastRow - 1
    
    For i = 2 To lastRow
        status = UCase(Trim(wsRR.Cells(i, 10).value))
        riskScore = CDbl(wsRR.Cells(i, 4).value)
        
        Select Case status
            Case "HOLD": holdCount = holdCount + 1
            Case "REVIEW": reviewCount = reviewCount + 1
            Case "PASS": passCount = passCount + 1
        End Select
        
        If riskScore >= 75 Then criticalCount = criticalCount + 1
        riskSum = riskSum + riskScore
    Next i
    
    avgRiskScore = riskSum / totalBatches
    
    summaryMsg = "RISK REGISTER SUMMARY" & vbCrLf & vbCrLf & _
                 "Total Batches: " & totalBatches & vbCrLf & _
                 "Critical (Score >= 75): " & criticalCount & vbCrLf & _
                 "Average Risk Score: " & Format(avgRiskScore, "0.0") & vbCrLf & vbCrLf & _
                 "STATUS DISTRIBUTION:" & vbCrLf & _
                 "  HOLD: " & holdCount & " (" & Format(holdCount / totalBatches, "0.0%") & ")" & vbCrLf & _
                 "  REVIEW: " & reviewCount & " (" & Format(reviewCount / totalBatches, "0.0%") & ")" & vbCrLf & _
                 "  PASS: " & passCount & " (" & Format(passCount / totalBatches, "0.0%") & ")"
    
    MsgBox summaryMsg, vbInformation, "Summary"
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub GenerateAnomalyReport()
    Dim wsRR As Worksheet, lastRow As Long, i As Long
    Dim reportFile As String, fileNum As Integer
    Dim anomalyCount As Long, riskScore As Double, confidence As Double
    Dim lineText As String
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    reportFile = ThisWorkbook.Path & "\AERPA_AnomalyReport_" & Format(Now(), "yyyymmddhhmmss") & ".txt"
    fileNum = FreeFile
    Open reportFile For Output As fileNum
    
    Print #fileNum, "AERPA ANOMALY DETECTION REPORT"
    Print #fileNum, "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, String(80, "=")
    Print #fileNum, ""
    
    For i = 2 To lastRow
        riskScore = CDbl(wsRR.Cells(i, 4).value)
        confidence = CDbl(wsRR.Cells(i, 5).value)
        
        If riskScore >= 60 And confidence < 0.5 Then
            Print #fileNum, "ANOMALY: " & wsRR.Cells(i, 2).value
            Print #fileNum, "  Risk: " & Format(riskScore, "0.0") & " | Confidence: " & Format(confidence, "0.0%")
            Print #fileNum, "  Drivers: " & wsRR.Cells(i, 6).value & ", " & wsRR.Cells(i, 7).value & ", " & wsRR.Cells(i, 8).value
            Print #fileNum, "  Status: " & wsRR.Cells(i, 10).value
            Print #fileNum, "  Equipment: " & wsRR.Cells(i, 13).value & " | Supplier: " & wsRR.Cells(i, 14).value
            Print #fileNum, ""
            anomalyCount = anomalyCount + 1
        End If
    Next i
    
    Print #fileNum, String(80, "=")
    Print #fileNum, "Total Anomalies: " & anomalyCount
    Close fileNum
    
    MsgBox "Report: " & reportFile & vbCrLf & "Anomalies: " & anomalyCount, vbInformation
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Close fileNum
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub AnalyzeDriverImpact()
    Dim wsRR As Worksheet, lastRow As Long, i As Long, j As Long
    Dim drivers As Object, driverRisks As Object, driverCounts As Object
    Dim avgRisk As Double, reportFile As String, fileNum As Integer
    Dim driver As Variant, riskScore As Double
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    Set drivers = CreateObject("Scripting.Dictionary")
    Set driverRisks = CreateObject("Scripting.Dictionary")
    Set driverCounts = CreateObject("Scripting.Dictionary")
    
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    For i = 2 To lastRow
        riskScore = CDbl(wsRR.Cells(i, 4).value)
        For j = 6 To 8   ' Driver1, Driver2, Driver3 (columns 6, 7, 8)
            driver = Trim(wsRR.Cells(i, j).value)
            If Len(driver) > 0 Then
                If drivers.Exists(driver) Then
                    drivers(driver) = drivers(driver) + 1
                    driverRisks(driver) = driverRisks(driver) + riskScore
                Else
                    drivers.Add driver, 1
                    driverRisks.Add driver, riskScore
                    driverCounts.Add driver, 0
                End If
            End If
        Next j
    Next i
    
    reportFile = ThisWorkbook.Path & "\AERPA_DriverAnalysis_" & Format(Now(), "yyyymmddhhmmss") & ".txt"
    fileNum = FreeFile
    Open reportFile For Output As fileNum
    
    Print #fileNum, "DRIVER IMPACT ANALYSIS"
    Print #fileNum, "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Print #fileNum, String(80, "=")
    Print #fileNum, ""
    
    For Each driver In drivers.keys
        avgRisk = driverRisks(driver) / drivers(driver)
        Print #fileNum, "Driver: " & driver
        Print #fileNum, "  Count: " & drivers(driver)
        Print #fileNum, "  Avg Risk: " & Format(avgRisk, "0.0")
        Print #fileNum, "  Total: " & Format(driverRisks(driver), "0.0")
        Print #fileNum, ""
    Next driver
    
    Close fileNum
    MsgBox "Report: " & reportFile & vbCrLf & "Drivers: " & drivers.count, vbInformation
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Close fileNum
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub GetTopRiskDrivers()
    Dim wsRR As Worksheet, lastRow As Long, i As Long, j As Long
    Dim drivers() As String, driverRisks() As Double, driverCounts() As Long
    Dim driverIndex As Long, found As Boolean, k As Long, temp As String, tempRisk As Double
    Dim riskScore As Double, driver As String, summaryMsg As String, tempCount As Long
    
    On Error GoTo ErrorHandler
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row
    
    If lastRow < 2 Then MsgBox "No data", vbExclamation: Exit Sub
    
    ReDim drivers(1 To 100)
    ReDim driverRisks(1 To 100)
    ReDim driverCounts(1 To 100)
    driverIndex = 0
    
    ' Collect unique drivers and their risk totals
    For i = 2 To lastRow
        riskScore = CDbl(wsRR.Cells(i, 4).value)
        For j = 6 To 8  ' Columns 6, 7, 8 (Driver1, Driver2, Driver3)
            driver = Trim(wsRR.Cells(i, j).value)
            If Len(driver) > 0 Then
                found = False
                For k = 1 To driverIndex
                    If drivers(k) = driver Then
                        driverRisks(k) = driverRisks(k) + riskScore
                        driverCounts(k) = driverCounts(k) + 1
                        found = True
                        Exit For
                    End If
                Next k
                If Not found Then
                    driverIndex = driverIndex + 1
                    drivers(driverIndex) = driver
                    driverRisks(driverIndex) = riskScore
                    driverCounts(driverIndex) = 1
                End If
            End If
        Next j
    Next i
    
    ' Sort by total risk (descending)
    Dim swapped As Boolean
    Do
        swapped = False
        For k = 1 To driverIndex - 1
            If driverRisks(k) < driverRisks(k + 1) Then
                tempRisk = driverRisks(k)
                driverRisks(k) = driverRisks(k + 1)
                driverRisks(k + 1) = tempRisk
                
                temp = drivers(k)
                drivers(k) = drivers(k + 1)
                drivers(k + 1) = temp
                
                tempCount = driverCounts(k)
                driverCounts(k) = driverCounts(k + 1)
                driverCounts(k + 1) = tempCount
                
                swapped = True
            End If
        Next k
    Loop While swapped
    
    ' Display top 3
    summaryMsg = "TOP RISK DRIVERS" & vbCrLf & vbCrLf
    Dim maxDrivers As Long
    maxDrivers = IIf(driverIndex < 3, driverIndex, 3)
    
    For k = 1 To maxDrivers
        summaryMsg = summaryMsg & k & ". " & drivers(k) & vbCrLf & _
                     "   Count: " & driverCounts(k) & _
                     " | Avg: " & Format(driverRisks(k) / driverCounts(k), "0.0") & _
                     " | Total: " & Format(driverRisks(k), "0.0") & vbCrLf & vbCrLf
    Next k
    
    MsgBox summaryMsg, vbInformation, "Top Risk Drivers"
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'================================================================================
' DATA ENTRY FORM
'================================================================================

Public Sub OpenRiskDataEntryForm()
    Dim wsForm As Worksheet
    On Error Resume Next
    Set wsForm = ThisWorkbook.Worksheets("RISK_ENTRY_FORM")
    On Error GoTo 0
    
    If wsForm Is Nothing Then Call CreateRiskDataEntryForm Else wsForm.Activate
    MsgBox "Risk Entry Form ready. Fill fields and click SUBMIT.", vbInformation
End Sub

Private Sub CreateRiskDataEntryForm()
    Dim wsForm As Worksheet, row As Long
    
    On Error Resume Next
    Set wsForm = ThisWorkbook.Worksheets("RISK_ENTRY_FORM")
    On Error GoTo 0
    
    If wsForm Is Nothing Then
        Set wsForm = ThisWorkbook.Sheets.Add
        wsForm.Name = "RISK_ENTRY_FORM"
    End If
    
    wsForm.Cells.Clear
    
    With wsForm.Range("A1:B1")
        .Merge
        .value = "RISK REGISTER DATA ENTRY FORM v10.2"
        .Interior.color = RGB(41, 84, 115)
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    
    row = 2
    Call AddFormField(wsForm, row, "Batch_ID *", "F_BatchID")
    Call AddFormField(wsForm, row, "Tenant_ID", "F_TenantID")
    Call AddFormField(wsForm, row, "Risk_Score (0-100) *", "F_RiskScore")
    Call AddFormField(wsForm, row, "Confidence (0-1.0) *", "F_Confidence")
    Call AddFormField(wsForm, row, "Driver1", "F_Driver1")
    Call AddFormField(wsForm, row, "Driver2", "F_Driver2")
    Call AddFormField(wsForm, row, "Driver3", "F_Driver3")
    Call AddFormField(wsForm, row, "Recommendation", "F_Recommendation")
    Call AddFormField(wsForm, row, "Status (PASS/REVIEW/HOLD)", "F_Status")
    Call AddFormField(wsForm, row, "Review_Notes", "F_ReviewNotes")
    Call AddFormField(wsForm, row, "Equipment_ID (EQP-xxxx)", "F_EquipmentID")
    Call AddFormField(wsForm, row, "Supplier_ID_Encoded (SUP-xxx)", "F_SupplierID")
    
    row = row + 1
    With wsForm.Range("A" & row & ":B" & row)
        .Merge
        .Interior.color = RGB(100, 200, 100)
        .value = "SUBMIT"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    wsForm.Activate
End Sub

Private Sub AddFormField(wsForm As Worksheet, ByRef row As Long, _
                         labelText As String, fieldName As String)
    ' Column A: label
    With wsForm.Range("A" & row)
        .value = labelText
        .Interior.color = RGB(41, 84, 115)
        .Font.color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    
    ' Column B: input cell
    With wsForm.Range("B" & row)
        .Name = fieldName
        .Interior.color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
    
    row = row + 1
End Sub

Public Sub SubmitRiskDataEntry()
    Dim wsForm As Worksheet, wsRR As Worksheet, lastRow As Long
    Dim batchID As String, tenantID As String, riskScore As Double
    Dim confidence As Double, recommendation As String
    Dim driver1 As String, driver2 As String, driver3 As String
    Dim status As String, reviewNotes As String
    Dim equipmentID As String, supplierID As String
    
    On Error GoTo ErrorHandler
    Set wsForm = ThisWorkbook.Worksheets("RISK_ENTRY_FORM")
    Set wsRR = ThisWorkbook.Worksheets("RISK_REGISTER")
    
    batchID = Trim(wsForm.Range("F_BatchID").value)
    tenantID = Trim(wsForm.Range("F_TenantID").value)
    riskScore = CDbl(wsForm.Range("F_RiskScore").value)
    confidence = CDbl(wsForm.Range("F_Confidence").value)
    driver1 = Trim(wsForm.Range("F_Driver1").value)
    driver2 = Trim(wsForm.Range("F_Driver2").value)
    driver3 = Trim(wsForm.Range("F_Driver3").value)
    recommendation = Trim(wsForm.Range("F_Recommendation").value)
    status = Trim(wsForm.Range("F_Status").value)
    reviewNotes = Trim(wsForm.Range("F_ReviewNotes").value)
    equipmentID = Trim(wsForm.Range("F_EquipmentID").value)
    supplierID = Trim(wsForm.Range("F_SupplierID").value)
    
    If Len(batchID) = 0 Then MsgBox "Batch_ID required", vbExclamation: Exit Sub
    If riskScore < 0 Or riskScore > 100 Then MsgBox "Risk_Score: 0-100", vbExclamation: Exit Sub
    If confidence < 0 Or confidence > 1 Then MsgBox "Confidence: 0-1", vbExclamation: Exit Sub
    
    lastRow = wsRR.Cells(wsRR.Rows.count, 2).End(xlUp).row + 1
    
    With wsRR
        .Cells(lastRow, 1).value = Now()              ' Timestamp
        .Cells(lastRow, 2).value = batchID           ' Batch_ID
        .Cells(lastRow, 3).value = tenantID          ' Tenant_ID
        .Cells(lastRow, 4).value = riskScore         ' Risk_Score
        .Cells(lastRow, 5).value = confidence        ' Confidence
        .Cells(lastRow, 6).value = driver1           ' Driver1
        .Cells(lastRow, 7).value = driver2           ' Driver2
        .Cells(lastRow, 8).value = driver3           ' Driver3
        .Cells(lastRow, 9).value = recommendation    ' Recommendation
        .Cells(lastRow, 10).value = status           ' Status
        .Cells(lastRow, 11).value = reviewNotes      ' Review_Notes
        .Cells(lastRow, 12).value = Environ("username") ' Reviewed_By
        .Cells(lastRow, 13).value = equipmentID      ' Equipment_ID
        .Cells(lastRow, 14).value = supplierID       ' Supplier_ID_Encoded
    End With
    
    wsForm.Range("F_BatchID:F_SupplierID").ClearContents
    
    MsgBox "Record submitted: " & batchID, vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'================================================================================
' END - AERPA v10.2 COMPLETE RISK REGISTER MODULE
'================================================================================


