VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repo_StockMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' STOCK MOVEMENT REPOSITORY (INVENTORY DATA LAYER)
' =========================================
' Responsible for persisting stock movement transactions
' into the StockMovements worksheet.
' Encapsulates write operations for inventory movements.

' Adds a new stock movement record
' Parameters:
' - mov: Stock movement entity containing inventory transaction data
Public Sub Add(mov As Ent_StockMovement)

    ' Validate input object
    If mov Is Nothing Then
        ERR.Raise vbObjectError + 30, "Repo_StockMovement.Add", "Movement is null"
    End If

    Dim ws As Worksheet
    
    ' Reference StockMovements worksheet
    Set ws = ThisWorkbook.Sheets("StockMovements")
    
    Dim lastRow As Long

    ' Handle case where sheet is empty (no headers or data)
    If ws.Cells(1, 1).value = "" Then
        lastRow = 1
    Else
        ' Determine next available row for insertion
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    End If
    
    ' Persist stock movement fields into worksheet
    ws.Cells(lastRow, 1).value = mov.Id
    ws.Cells(lastRow, 2).value = mov.ProductId
    ws.Cells(lastRow, 3).value = mov.Quantity
    ws.Cells(lastRow, 4).value = mov.MovementType
    ws.Cells(lastRow, 5).value = mov.DocumentId
    ws.Cells(lastRow, 6).value = mov.CreatedAt

End Sub
