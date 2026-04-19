VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repo_ProcessedOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' PROCESSED OPERATIONS REPOSITORY
' =========================================
' Manages idempotent operation tracking by storing processed operations.
' Ensures duplicate operations are prevented using CorrelationId and OperationType.

' =========================
' SAFE SHEET ACCESS
' =========================
' Retrieves worksheet by name in a safe, case-insensitive manner
Private Function GetSheetSafe(sheetName As String) As Worksheet

    Dim ws As Worksheet
    
    ' Iterate through all worksheets to find a match
    For Each ws In ThisWorkbook.Worksheets
        If Trim(LCase(ws.Name)) = Trim(LCase(sheetName)) Then
            Set GetSheetSafe = ws
            Exit Function
        End If
    Next ws
    
    ' Raise error if sheet is not found
    ERR.Raise 999, , "Sheet not found: " & sheetName

End Function


' =========================
' NORMALIZE HELPER
' =========================
' Normalizes string values for consistent comparison
' (trims whitespace and converts to lowercase)
Private Function Normalize(value As String) As String
    Normalize = Trim(LCase(value))
End Function


' =========================
' EXISTS CHECK (SAFE)
' =========================
' Checks whether a processed operation already exists
' based on CorrelationId and OperationType
Public Function Exists(CorrelationId As String, OperationType As String) As Boolean

    Dim ws As Worksheet
    Set ws = GetSheetSafe("ProcessedOperations")
    
    Dim lastRow As Long
    
    ' Determine last used row
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Exit if no data rows exist
    If lastRow < 2 Then Exit Function
    
    Dim i As Long
    
    ' Iterate through stored operations
    For i = 2 To lastRow
        
        ' Compare normalized values for idempotency check
        If Normalize(ws.Cells(i, 1).value) = Normalize(CorrelationId) _
        And Normalize(ws.Cells(i, 2).value) = Normalize(OperationType) Then
            
            Exists = True
            Exit Function
            
        End If
        
    Next i

End Function


' =========================
' ADD WITH UNIQUE CONTROL
' =========================
' Adds a processed operation record with uniqueness enforcement
' Prevents duplicate processing using CorrelationId + OperationType
Public Sub Add(CorrelationId As String, OperationType As String, EntityId As Long)

    Dim ws As Worksheet
    Set ws = GetSheetSafe("ProcessedOperations")
    
    ' Enforce uniqueness (critical for idempotency)
    If Exists(CorrelationId, OperationType) Then
        ERR.Raise vbObjectError + 600, "Idempotency", _
        "Duplicate operation blocked: " & CorrelationId & " / " & OperationType
    End If
    
    Dim lastRow As Long
    
    ' Determine insertion row
    If ws.Cells(1, 1).value = "" Then
        lastRow = 1
    Else
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    End If
    
    ' Persist processed operation record
    ws.Cells(lastRow, 1).value = CorrelationId
    ws.Cells(lastRow, 2).value = OperationType
    ws.Cells(lastRow, 3).value = EntityId
    ws.Cells(lastRow, 4).value = "Completed"
    ws.Cells(lastRow, 5).value = Now

End Sub

