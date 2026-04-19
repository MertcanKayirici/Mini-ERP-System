VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repo_Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' AUDIT LOGGING SERVICE (AUDIT TRAIL LAYER)
' =========================================
' Responsible for recording system actions into the AuditLog worksheet.
' Provides traceability and supports debugging, monitoring, and compliance.

' Logs an action performed in the system
' Parameters:
' - Action: Describes the operation (e.g., CREATE, UPDATE, DELETE, POST)
' - EntityId: Identifier of the affected entity
' - EntityType: Type of the entity (e.g., Product, Document)
Public Sub LogAction(Action As String, EntityId As Long, EntityType As String)

    Dim ws As Worksheet
    
    ' Reference AuditLog worksheet
    Set ws = ThisWorkbook.Sheets("AuditLog")
    
    Dim lastRow As Long
    
    ' Determine next available row for insertion
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ' Write audit log entry fields
    ws.Cells(lastRow, 1).value = Action
    ws.Cells(lastRow, 2).value = EntityId
    ws.Cells(lastRow, 3).value = EntityType
    
    ' Capture system user executing the action
    ws.Cells(lastRow, 4).value = Environ("Username")
    
    ' Capture timestamp of the action
    ws.Cells(lastRow, 5).value = Now

End Sub
