VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Qry_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' STOCK QUERY (SINGLE PRODUCT)
' =========================================
' Retrieves current stock for a specific product
' via service layer with validation and safe fallback.

Public Function GetStock(ProductId As Long) As Double

    ' Initialize stock service
    Dim stockSvc As New Svc_Stock
    
    ' Safe execution block
    On Error GoTo FAIL
    
    ' Retrieve current stock from service
    GetStock = stockSvc.GetCurrentStock(ProductId)
    
    ' Validate data integrity (stock should not be negative)
    If GetStock < 0 Then
        ERR.Raise 2001, "Qry_Stock", "Stock cannot be negative"
    End If
    
    Exit Function

FAIL:
    ' Fallback value in case of error
    GetStock = 0

End Function


' =========================================
' STOCK QUERY (TOTAL STOCK)
' =========================================
' Calculates total stock across all products
' by aggregating StockMovements sheet.

Public Function GetTotalStock() As Double

    ' Reference StockMovements worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("StockMovements")
    
    ' Determine last used row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim total As Double
    Dim i As Long
    
    ' Iterate through all stock movement records
    For i = 2 To lastRow
        
        ' Apply movement logic based on type
        Select Case ws.Cells(i, 4).value
            Case "IN"
                total = total + ws.Cells(i, 3).value   ' Increase stock
            Case "OUT"
                total = total - ws.Cells(i, 3).value   ' Decrease stock
        End Select
        
    Next i
    
    ' Return aggregated total stock
    GetTotalStock = total

End Function
