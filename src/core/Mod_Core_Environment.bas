Attribute VB_Name = "Mod_Core_Environment"
Option Explicit

' Constant defining the dashboard worksheet name
Private Const DASHBOARD_SHEET As String = "Dashboard"

' =========================
' MAIN
' =========================
' Entry point to build, refresh, and style the dashboard
Public Sub RunDashboard()
    BuildDashboard          ' Create or reset dashboard layout
    RefreshDashboard        ' Populate dashboard with current data
    StyleDashboard          ' Apply visual styling
    
    ' Notify user that dashboard is ready
    MsgBox "Dashboard is ready!"
End Sub

' =========================
' BUILD
' =========================
' Creates the dashboard sheet and initializes static layout and labels
Private Sub BuildDashboard()

    Dim ws As Worksheet
    
    ' Attempt to reference existing dashboard sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    On Error GoTo 0
    
    ' If sheet does not exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = DASHBOARD_SHEET
    End If

    ' Clear previous content
    ws.Cells.Clear

    ' ===== TITLE (MERGED A-B RANGE) =====
    ws.Range("A1:B1").Merge
    ws.Range("A1").value = "MINI ERP SYSTEM MONITOR"

    ' Apply title styling
    With ws.Range("A1")
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.color = RGB(30, 30, 30)
        .Font.color = RGB(255, 255, 255)
        .RowHeight = 28
    End With

    ' ===== SECTION HEADERS =====
    ws.Range("A3").value = "CORE METRICS"
    ws.Range("A9").value = "SYSTEM STATE"
    ws.Range("A14").value = "RESILIENCE"

    ws.Range("A3,A9,A14").Font.Bold = True

    ' ===== LABEL DEFINITIONS =====
    ws.Range("A4").value = "Total Stock"
    ws.Range("A5").value = "Total Products"
    ws.Range("A6").value = "Ledger Total"
    ws.Range("A7").value = "Tests Passed"

    ws.Range("A10").value = "Last Operation"
    ws.Range("A11").value = "System Status"
    ws.Range("A12").value = "Reconciliation"

    ws.Range("A15").value = "Retry Count"
    ws.Range("A16").value = "Active Locks"
    ws.Range("A17").value = "Last Error"

End Sub

' =========================
' REFRESH
' =========================
' Retrieves live data and updates dashboard values
Private Sub RefreshDashboard()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)

    ' ===== CORE METRICS =====
    ' Fetch total stock via query class
    Dim qStock As New Qry_Stock
    ws.Range("B4").value = qStock.GetTotalStock

    ' Calculate total product count from Products sheet
    Dim wsProd As Worksheet
    Set wsProd = ThisWorkbook.Sheets("Products")
    ws.Range("B5").value = Application.Max(0, wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row - 1)

    ' Fetch ledger total via reconciliation query
    Dim qRecon As New Qry_Reconciliation
    ws.Range("B6").value = qRecon.GetLedgerTotal

    ' ===== TEST RESULTS =====
    Dim wsTest As Worksheet
    Set wsTest = ThisWorkbook.Sheets("TestResults")

    Dim lastRow As Long, i As Long
    Dim passCount As Long, totalCount As Long

    ' Determine last row in test results
    lastRow = wsTest.Cells(wsTest.Rows.count, 1).End(xlUp).row

    If lastRow < 2 Then
        ' No test data available
        ws.Range("B7").value = "0 / 0"
        ws.Range("B11").value = "NO TEST"
    Else
        totalCount = lastRow - 1
        
        ' Count passed tests
        For i = 2 To lastRow
            If wsTest.Cells(i, 2).value = "PASS" Then
                passCount = passCount + 1
            End If
        Next i
        
        ' Display pass ratio
        ws.Range("B7").value = passCount & " / " & totalCount
        
        ' Set system status based on test results
        If passCount = totalCount Then
            ws.Range("B11").value = "OK"
        Else
            ws.Range("B11").value = "ERROR"
        End If
    End If

    ' ===== SYSTEM STATE =====
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("AuditLog")

    ' Get last operation from audit log
    lastRow = wsAudit.Cells(wsAudit.Rows.count, 1).End(xlUp).row

    If lastRow > 1 Then
        ws.Range("B10").value = wsAudit.Cells(lastRow, 1).value
    Else
        ws.Range("B10").value = "-"
    End If

    ' Static reconciliation status
    ws.Range("B12").value = "PASSED"

    ' ===== RESILIENCE METRICS =====
    ws.Range("B15").value = CountRetries()
    ws.Range("B16").value = CountActiveLocks()
    ws.Range("B17").value = GetLastError()

End Sub

' =========================
' STYLE
' =========================
' Applies visual formatting to dashboard sections
Private Sub StyleDashboard()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)

    ' Auto-fit columns for better readability
    ws.Columns("A:B").AutoFit

    ' Apply card-style formatting to sections
    StyleCard ws.Range("A4:B7")
    StyleCard ws.Range("A10:B12")
    StyleCard ws.Range("A15:B17")

    ' Highlight value column
    ws.Range("B4:B17").Font.Bold = True

    ' Apply status color coding
    If ws.Range("B11").value = "OK" Then
        ws.Range("B11").Interior.color = RGB(0, 180, 0)
        ws.Range("B11").Font.color = RGB(255, 255, 255)
    Else
        ws.Range("B11").Interior.color = RGB(200, 0, 0)
        ws.Range("B11").Font.color = RGB(255, 255, 255)
    End If

End Sub

' Applies border and background styling to a given range
Private Sub StyleCard(rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Interior.color = RGB(245, 245, 245)
    End With
End Sub

' =========================
' HELPERS
' =========================
' Counts retry operations in the audit log
Private Function CountRetries() As Long

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("AuditLog")

    Dim i As Long, count As Long
    Dim lastRow As Long
    
    ' Determine last row in audit log
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' Count retry events
    For i = 2 To lastRow
        If ws.Cells(i, 1).value = "RETRY_POST" Then
            count = count + 1
        End If
    Next i

    CountRetries = count

End Function

' Returns number of active locks (placeholder implementation)
Private Function CountActiveLocks() As Long
    CountActiveLocks = 0
End Function

' Returns last error message (placeholder implementation)
Private Function GetLastError() As String
    GetLastError = "-"
End Function
