VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Svc_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' STOCK SERVICE (INVENTORY BUSINESS LOGIC LAYER)
' =========================================
' Handles stock-related operations including:
' - Adding stock (IN movements)
' - Removing stock (OUT movements)
' - Calculating stock levels
' Uses StockMovements as an append-only event log.

' =========================
' ADD STOCK (IN)
' =========================
' Adds stock for a given product
' - Validates input quantity
' - Creates a stock movement record (IN)
' - Persists movement via repository
Public Sub AddStock(ProductId As Long, Quantity As Double)

    ' Validate quantity input
    If Quantity <= 0 Then
        ERR.Raise vbObjectError + 1, "Svc_Stock.AddStock", "Quantity must be positive"
    End If

    Dim mov As New Ent_StockMovement
    
    ' Initialize movement entity
    mov.Id = GenerateId()
    mov.ProductId = ProductId
    mov.Quantity = Quantity
    mov.MovementType = "IN"
    mov.DocumentId = 0
    mov.CreatedAt = Now

    Dim repo As New Repo_StockMovement
    
    ' Persist stock movement
    repo.Add mov

End Sub


' =========================
' REMOVE STOCK (OUT)
' =========================
' Removes stock for a given product
' - Validates input quantity
' - Ensures sufficient stock
' - Creates a stock movement record (OUT)
Public Sub RemoveStock(ProductId As Long, Quantity As Double)

    ' Validate quantity input
    If Quantity <= 0 Then
        ERR.Raise vbObjectError + 1, "Svc_Stock.RemoveStock", "Quantity must be positive"
    End If

    Dim currentStock As Double
    
    ' Retrieve current stock level
    currentStock = GetCurrentStock(ProductId)

    ' Prevent negative stock scenarios
    If currentStock < Quantity Then
        ERR.Raise vbObjectError + 2, "Svc_Stock.RemoveStock", "Insufficient stock"
    End If

    Dim mov As New Ent_StockMovement
    
    ' Initialize movement entity
    mov.Id = GenerateId()
    mov.ProductId = ProductId
    mov.Quantity = Quantity
    mov.MovementType = "OUT"
    mov.DocumentId = 0
    mov.CreatedAt = Now

    Dim repo As New Repo_StockMovement
    
    ' Persist stock movement
    repo.Add mov

End Sub


' =========================
' GET CURRENT STOCK
' =========================
' Calculates current stock level for a specific product
' based on StockMovements event log.
Public Function GetCurrentStock(ProductId As Long) As Double

    Dim ws As Worksheet
    
    ' Reference StockMovements worksheet
    Set ws = ThisWorkbook.Sheets("StockMovements")
    
    Dim lastRow As Long
    
    ' Determine last used row
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Return zero if no data exists
    If lastRow < 2 Then
        GetCurrentStock = 0
        Exit Function
    End If
    
    Dim i As Long
    Dim total As Double
    
    ' Aggregate stock movements
    For i = 2 To lastRow
        
        ' Filter by ProductId
        If ws.Cells(i, 2).value = ProductId Then
            
            Select Case ws.Cells(i, 4).value
                
                Case "IN"
                    total = total + ws.Cells(i, 3).value   ' Increase stock
                    
                Case "OUT"
                    total = total - ws.Cells(i, 3).value   ' Decrease stock
                    
            End Select
            
        End If
        
    Next i
    
    GetCurrentStock = total

End Function


' =========================
' ID GENERATOR
' =========================
' Generates incremental unique ID (in-memory)
' NOTE:
' - Resets on application restart
' - Suitable for local/demo scenarios
Private Function GenerateId() As Long

    ' Static variable persists within session
    Static currentId As Long
    
    ' Increment ID counter
    currentId = currentId + 1
    
    GenerateId = currentId

End Function


' =========================
' GET TOTAL STOCK
' =========================
' Calculates total stock across all products
' by aggregating all movement records
Public Function GetTotalStock() As Double

    Dim ws As Worksheet
    
    ' Reference StockMovements worksheet
    Set ws = ThisWorkbook.Sheets("StockMovements")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim total As Double
    Dim i As Long
    
    ' Iterate through all movement records
    For i = 2 To lastRow
        
        Select Case ws.Cells(i, 4).value
            
            Case "IN"
                total = total + ws.Cells(i, 3).value   ' Increase total
                
            Case "OUT"
                total = total - ws.Cells(i, 3).value   ' Decrease total
                
        End Select
        
    Next i
    
    GetTotalStock = total

End Function
