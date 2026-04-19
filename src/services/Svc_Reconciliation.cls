VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Svc_Reconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' RECONCILIATION SERVICE (DATA CONSISTENCY CHECK)
' =========================================
' Validates system integrity by comparing calculated stock
' from movement records with system-reported stock values.

Public Sub RunFullReconciliation()

    On Error GoTo FAIL

    ' Execute reconciliation for a single product (default test scope)
    ValidateStock 1

    ' Notify success
    MsgBox "RECONCILIATION SUCCESSFUL"
    Exit Sub

FAIL:
    ' Notify failure with error details
    MsgBox "RECONCILIATION FAILED: " & ERR.Description

End Sub


' Validates stock consistency for a given product
' Compares calculated stock (from movements) vs system stock (service)
Public Sub ValidateStock(ProductId As Long)

    Dim ws As Worksheet
    
    ' Reference StockMovements worksheet
    Set ws = ThisWorkbook.Sheets("StockMovements")
    
    Dim lastRow As Long
    
    ' Determine last used row
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    Dim calculatedStock As Double
    
    ' Calculate stock based on movement records
    For i = 2 To lastRow
        
        ' Filter by ProductId
        If ws.Cells(i, 2).value = ProductId Then
            
            ' Apply movement logic
            Select Case ws.Cells(i, 4).value
                Case "IN"
                    calculatedStock = calculatedStock + ws.Cells(i, 3).value
                    
                Case "OUT"
                    calculatedStock = calculatedStock - ws.Cells(i, 3).value
            End Select
            
        End If
        
    Next i
    
    ' Retrieve system-calculated stock
    Dim stockSvc As Svc_Stock
    Set stockSvc = New Svc_Stock
    
    Dim systemStock As Double
    systemStock = stockSvc.GetCurrentStock(ProductId)
    
    ' Compare calculated vs system values with tolerance
    If Abs(calculatedStock - systemStock) > 0.0001 Then
        
        ERR.Raise vbObjectError + 500, "Reconciliation", _
        "Stock mismatch! Calculated=" & calculatedStock & " System=" & systemStock
        
    End If

End Sub

