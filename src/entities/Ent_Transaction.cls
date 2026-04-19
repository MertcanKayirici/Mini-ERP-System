VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ent_Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' TRANSACTION CONTEXT (ROLLBACK MANAGER)
' =========================================
' Maintains a collection of rollback actions to support
' transactional consistency. Each action represents a step
' that can be reverted in case of failure.

' Collection storing rollback actions
' Each item is an array: [actionName, data]
Public RollbackActions As Collection

' Initializes the rollback collection when class instance is created
Private Sub Class_Initialize()
    Set RollbackActions = New Collection
End Sub

' Adds a rollback action to the collection
' actionName: identifier of the rollback operation
' data: context data required to execute the rollback
Public Sub AddRollback(actionName As String, data As Variant)

    ' Store action metadata as a 2-element array
    Dim item(1) As Variant
    item(0) = actionName
    item(1) = data
    
    ' Append rollback action to collection
    RollbackActions.Add item

End Sub
