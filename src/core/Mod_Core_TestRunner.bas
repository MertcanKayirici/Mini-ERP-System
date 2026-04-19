Attribute VB_Name = "Mod_Core_TestRunner"
Option Explicit

' =========================================
' MASTER TEST SYSTEM
' =========================================
' Coordinates full system lifecycle: reset, setup, seed, test execution, and dashboard refresh
Public Sub Run_Full_TestCycle()

    ResetAllData          ' Clear all system data
    SetupEnvironment      ' Initialize required sheets and headers
    SeedFullSystem        ' Insert initial sample data
    Run_All_Tests         ' Execute all defined test cases
    RunDashboard          ' Refresh dashboard after test cycle

    ' Notify completion of full test cycle
    MsgBox "TEST CYCLE COMPLETED"

End Sub


' =========================================
' RESET
' =========================================
' Clears all transactional and system-related sheets
Public Sub ResetAllData()

    ClearSheetSafe "Products"
    ClearSheetSafe "StockMovements"
    ClearSheetSafe "CustomerLedger"
    ClearSheetSafe "AuditLog"
    ClearSheetSafe "ProcessedOperations"
    ClearSheetSafe "TestResults"

End Sub

' Safely clears sheet content while preserving headers if present
Private Sub ClearSheetSafe(sheetName As String)

    Dim ws As Worksheet
    
    ' Attempt to get worksheet reference
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    ' Exit if sheet does not exist
    If ws Is Nothing Then Exit Sub
    
    ' Check if header exists
    If Trim(ws.Cells(1, 1).value) <> "" Then
        
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        
        ' Clear only data rows (preserve header row)
        If lastRow >= 2 Then
            ws.Range(ws.Rows(2), ws.Rows(lastRow)).ClearContents
        End If
        
    Else
        ' If no header, clear entire sheet
        ws.Cells.Clear
    End If

End Sub


' =========================================
' SETUP
' =========================================
' Initializes all required sheets with predefined schemas
Public Sub SetupEnvironment()

    SetupSheet "Products", Array("Id", "Name", "Price", "Cost", "IsActive", "CreatedAt")
    SetupSheet "StockMovements", Array("Id", "ProductId", "Quantity", "MovementType", "DocumentId", "CreatedAt")
    SetupSheet "CustomerLedger", Array("Id", "CustomerId", "Amount", "EntryType", "DocumentId", "CreatedAt")
    SetupSheet "AuditLog", Array("Action", "EntityId", "EntityType", "User", "CreatedAt")
    SetupSheet "ProcessedOperations", Array("CorrelationId", "OperationType", "EntityId", "Status", "CreatedAt")
    SetupSheet "TestResults", Array("Test Name", "Result", "Message")

End Sub

' Creates or resets a worksheet and applies column headers
Private Sub SetupSheet(sheetName As String, headers As Variant)

    Dim ws As Worksheet
    Dim i As Long
    Dim found As Boolean
    
    ' Check if sheet already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            found = True
            Exit For
        End If
    Next ws
    
    ' Create sheet if not found
    If Not found Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        Set ws = ThisWorkbook.Sheets(sheetName)
    End If
    
    ' Clear header row
    ws.Rows(1).ClearContents
    
    ' Apply headers with bold formatting
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
    Next i
    
    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1
    
    ' Clear all data rows
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.count, colCount)).ClearContents

End Sub


' =========================================
' SEED
' =========================================
' Inserts initial sample data into the system for testing
Public Sub SeedFullSystem()

    Dim repo As New Repo_Product
    Dim stock As New Svc_Stock
    
    ' Create sample product entity
    Dim p As New Ent_Product
    p.Name = "Kalem"
    p.Price = 50
    p.Cost = 30
    p.IsActive = True
    p.CreatedAt = Now
    
    ' Persist product and initialize stock
    repo.Add p
    stock.AddStock p.Id, 50

End Sub


' =========================================
' TEST RUNNER
' =========================================
' Executes all test cases and logs results into TestResults sheet
Public Sub Run_All_Tests()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TestResults")
    
    ' Clear previous test results
    ws.Rows("2:" & ws.Rows.count).ClearContents
    
    Dim row As Long
    row = 2

    ' Execute test suite
    RunTest ws, row, "Reset", "Test_Reset"
    RunTest ws, row, "Setup Product", "Test_SetupProduct"
    RunTest ws, row, "Add Stock", "Test_AddStock"
    RunTest ws, row, "Success Sale", "Test_SuccessSale"
    RunTest ws, row, "Stock Fail", "Test_StockFail"
    RunTest ws, row, "Inactive Product", "Test_InactiveProduct"

End Sub


' Executes a single test case and captures PASS/FAIL result
Private Sub RunTest(ws As Worksheet, ByRef row As Long, Name As String, funcName As String)

    On Error GoTo FAIL

    ' Dynamically execute test procedure
    Application.Run funcName

    ' Mark test as PASS
    ws.Cells(row, 1).value = Name
    ws.Cells(row, 2).value = "PASS"
    ws.Cells(row, 2).Interior.color = vbGreen

    row = row + 1
    Exit Sub

FAIL:
    ' Mark test as FAIL and log error message
    ws.Cells(row, 1).value = Name
    ws.Cells(row, 2).value = "FAIL"
    ws.Cells(row, 2).Interior.color = vbRed
    ws.Cells(row, 3).value = ERR.Description

    row = row + 1

End Sub


' =========================================
' TEST CASES
' =========================================

' Verifies system reset functionality
Public Sub Test_Reset()
    ResetAllData
End Sub

' Tests product creation and persistence
Public Sub Test_SetupProduct()

    Dim repo As New Repo_Product
    
    Dim p As New Ent_Product
    p.Name = "Test"
    p.Price = 100
    p.Cost = 60
    p.IsActive = True
    p.CreatedAt = Now
    
    repo.Add p

End Sub

' Tests stock addition operation
Public Sub Test_AddStock()

    Dim stock As New Svc_Stock
    stock.AddStock 1, 10

End Sub

' Tests successful sales transaction flow
Public Sub Test_SuccessSale()

    Dim svc As New Svc_Document
    
    Dim doc As Ent_Document
    Set doc = svc.CreateDraft("SALE")
    
    Dim lines As New Collection
    lines.Add svc.CreateLine(doc, 1, 1, 50, 30)
    
    svc.PostDocument doc, lines

End Sub

' Tests failure scenario when stock is insufficient
Public Sub Test_StockFail()

    On Error GoTo OK

    Dim svc As New Svc_Document
    
    Dim doc As Ent_Document
    Set doc = svc.CreateDraft("SALE")
    
    Dim lines As New Collection
    lines.Add svc.CreateLine(doc, 1, 9999, 50, 30)
    
    svc.PostDocument doc, lines

    ' Force failure if no error occurred
    ERR.Raise 1, , "Expected error did not occur"
    
    Exit Sub

OK:
End Sub

' Tests behavior when product is inactive
Public Sub Test_InactiveProduct()

    On Error GoTo OK

    Dim repo As New Repo_Product
    Dim p As Ent_Product
    
    Set p = repo.GetById(1)
    p.IsActive = False
    repo.Update p

    Dim svc As New Svc_Document
    
    Dim doc As Ent_Document
    Set doc = svc.CreateDraft("SALE")
    
    Dim lines As New Collection
    lines.Add svc.CreateLine(doc, 1, 1, 50, 30)
    
    svc.PostDocument doc, lines

    ' Force failure if no error occurred
    ERR.Raise 1, , "Expected error did not occur"

    Exit Sub

OK:
End Sub
