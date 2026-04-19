VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Qry_Reconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' STOCK QUERY (FROM MOVEMENTS)
' =========================================
' Calculates current stock for a given product
' by iterating through StockMovements sheet.
' Aggregates IN and OUT movements.

Public Function GetStockFromMovements(ProductId As Long) As Double

    ' Reference StockMovements worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("StockMovements")
    
    ' Determine last used row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim total As Double
    Dim i As Long
    
    ' Iterate through movement records (skip header row)
    For i = 2 To lastRow
        
        ' Filter by ProductId
        If ws.Cells(i, 2).value = ProductId Then
            
            ' Apply movement logic based on type
            Select Case ws.Cells(i, 4).value
                Case "IN"
                    total = total + ws.Cells(i, 3).value   ' Increase stock
                Case "OUT"
                    total = total - ws.Cells(i, 3).value   ' Decrease stock
            End Select
            
        End If
        
    Next i
    
    ' Return calculated stock
    GetStockFromMovements = total

End Function


' =========================================
' LEDGER QUERY (TOTAL BALANCE)
' =========================================
' Calculates total ledger balance by summing all
' transaction amounts from CustomerLedger sheet.

Public Function GetLedgerTotal() As Double

    ' Reference CustomerLedger worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CustomerLedger")
    
    ' Determine last used row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim total As Double
    Dim i As Long
    
    ' Iterate through ledger entries (skip header row)
    For i = 2 To lastRow
        total = total + ws.Cells(i, 3).value
    Next i
    
    ' Return aggregated ledger total
    GetLedgerTotal = total

End Function
