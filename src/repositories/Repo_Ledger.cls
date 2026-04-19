VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repo_Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' CUSTOMER LEDGER REPOSITORY (DATA ACCESS LAYER)
' =========================================
' Responsible for persisting ledger entries into the CustomerLedger worksheet.
' Encapsulates data access logic for ledger operations.

' Adds a new ledger entry to the system
' Parameters:
' - entry: Ledger entity containing transaction data
Public Sub Add(entry As Ent_Ledger)

    Dim ws As Worksheet
    
    ' Retrieve worksheet using safe lookup method
    Set ws = GetSheetSafe("CustomerLedger")
    
    Dim lastRow As Long
    
    ' Determine next available row for insertion
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ' Map entity fields to worksheet columns
    ws.Cells(lastRow, 1).value = entry.Id
    ws.Cells(lastRow, 2).value = entry.CustomerId
    ws.Cells(lastRow, 3).value = entry.Amount
    ws.Cells(lastRow, 4).value = entry.EntryType
    ws.Cells(lastRow, 5).value = entry.DocumentId
    ws.Cells(lastRow, 6).value = entry.CreatedAt

End Sub

' Safely retrieves a worksheet by name (case-insensitive)
' Throws error if sheet is not found
Private Function GetSheetSafe(sheetName As String) As Worksheet

    Dim ws As Worksheet
    
    ' Iterate through all worksheets to find match
    For Each ws In ThisWorkbook.Worksheets
        If Trim(LCase(ws.Name)) = Trim(LCase(sheetName)) Then
            Set GetSheetSafe = ws
            Exit Function
        End If
    Next ws
    
    ' Raise error if sheet does not exist
    ERR.Raise 999, , "Sheet not found: " & sheetName

End Function
