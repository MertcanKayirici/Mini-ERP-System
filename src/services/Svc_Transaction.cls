VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Svc_Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' TRANSACTION SERVICE (ROLLBACK HANDLER)
' =========================================
' Executes rollback operations in reverse order
' to restore system consistency in case of failure.

Public Sub Rollback(tx As Ent_Transaction)

    Dim i As Long
    Dim item As Variant
    
    ' Execute rollback actions in reverse order (LIFO)
    For i = tx.RollbackActions.count To 1 Step -1
        
        item = tx.RollbackActions(i)
        
        ' Dispatch rollback operation based on action type
        Select Case item(0)
        
            Case "ADD_STOCK"
                UndoAddStock item(1)
                
            Case "REMOVE_STOCK"
                UndoRemoveStock item(1)
                
        End Select
        
    Next i

End Sub


' Reverts a previous stock addition by removing stock
Private Sub UndoAddStock(data As Variant)

    Dim stock As New Svc_Stock
    
    ' data(0): ProductId, data(1): Quantity
    stock.RemoveStock data(0), data(1)

End Sub


' Reverts a previous stock removal by adding stock back
Private Sub UndoRemoveStock(data As Variant)

    Dim stock As New Svc_Stock
    
    ' data(0): ProductId, data(1): Quantity
    stock.AddStock data(0), data(1)

End Sub
