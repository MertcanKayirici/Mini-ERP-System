VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Svc_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' DOCUMENT SERVICE (APPLICATION LAYER)
' =========================================
' Handles document lifecycle operations including creation,
' posting, cancellation, and integration with stock, ledger,
' audit, and idempotency systems.

' =========================
' CREATE DRAFT
' =========================
' Initializes a new draft document with default values
Public Function CreateDraft(DocumentType As String) As Ent_Document

    Dim doc As Ent_Document
    Dim audit As Repo_Audit

    ' Initialize dependencies
    Set doc = New Ent_Document
    Set audit = New Repo_Audit
    
    ' Assign core properties
    doc.Id = GenerateId()
    doc.DocumentType = DocumentType
    doc.Status = "Draft"
    doc.CorrelationId = GenerateCorrelationId()
    doc.IsLocked = False
    
    ' Initialize integration flags
    doc.StockCreated = False
    doc.LedgerCreated = False
    doc.AuditCreated = False
    
    ' Log creation event
    audit.LogAction "CREATE", doc.Id, "Document"
    
    Set CreateDraft = doc

End Function

' =========================
' CREATE LINE
' =========================
' Creates a document line associated with a parent document
Public Function CreateLine(doc As Ent_Document, ProductId As Long, Quantity As Double, Price As Double, Cost As Double) As Ent_DocumentLine

    Dim line As Ent_DocumentLine
    Set line = New Ent_DocumentLine
    
    ' Assign line properties
    line.Id = GenerateId()
    line.DocumentId = doc.Id
    line.ProductId = ProductId
    line.Quantity = Quantity
    line.Price = Price
    line.Cost = Cost
    
    Set CreateLine = line

End Function

' =========================
' POST DOCUMENT
' =========================
' Processes a document and applies business logic:
' - Stock operations
' - Ledger entries
' - Audit logging
' - Idempotency control
' - Transaction rollback on failure
Public Sub PostDocument(doc As Ent_Document, lines As Collection)

    On Error GoTo ERROR_HANDLER

    ' Initialize dependencies
    Dim tx As Ent_Transaction
    Dim txSvc As Svc_Transaction
    Dim opRepo As Repo_ProcessedOperations
    Dim audit As Repo_Audit
    Dim policy As Pol_Document
    Dim stockSvc As Svc_Stock
    Dim repoLedger As Repo_Ledger

    Set tx = New Ent_Transaction
    Set txSvc = New Svc_Transaction
    Set opRepo = New Repo_ProcessedOperations
    Set audit = New Repo_Audit
    Set policy = New Pol_Document
    Set stockSvc = New Svc_Stock
    Set repoLedger = New Repo_Ledger

    ' Validate document state
    policy.CheckLock doc
    policy.CanPost doc

    ' Validate input lines
    If lines Is Nothing Or lines.count = 0 Then
        ERR.Raise vbObjectError + 15, , "No lines to process"
    End If

    ' Lock document for processing
    doc.Status = "Posting"
    doc.IsLocked = True
    doc.LockedAt = Now

    Dim i As Long
    Dim line As Ent_DocumentLine
    Dim revenue As Double

    ' Process each document line
    For i = 1 To lines.count
        
        Set line = lines(i)
        
        ' Validate line existence
        If line Is Nothing Then
            ERR.Raise vbObjectError + 20, , "Line is empty"
        End If

        ' Ensure product is active
        policy.CheckProductActive line.ProductId

        ' Execute logic based on document type
        Select Case doc.DocumentType

            Case "STOCK_IN"
                ' Increase stock
                stockSvc.AddStock line.ProductId, line.Quantity
                
                ' Register rollback action
                tx.AddRollback "ADD_STOCK", Array(line.ProductId, line.Quantity)
                
            Case "SALE"
                ' Validate stock availability
                policy.CheckStockAvailability line.ProductId, line.Quantity
                
                ' Decrease stock
                stockSvc.RemoveStock line.ProductId, line.Quantity
                
                ' Register rollback action
                tx.AddRollback "REMOVE_STOCK", Array(line.ProductId, line.Quantity)
                
                ' Accumulate revenue
                revenue = revenue + (line.Quantity * line.Price)

            Case Else
                ' Unsupported document type
                ERR.Raise vbObjectError + 40, , "Unknown document type"
        
        End Select
        
    Next i

    ' Create ledger entry for sales
    If doc.DocumentType = "SALE" Then
    
        Dim ledger As Ent_Ledger
        Set ledger = New Ent_Ledger
        
        ledger.Id = GenerateId()
        ledger.CustomerId = 1
        ledger.Amount = revenue
        ledger.EntryType = "SALE"
        ledger.DocumentId = doc.Id
        ledger.CreatedAt = Now
        
        repoLedger.Add ledger
        
    End If

    ' Log posting action
    audit.LogAction "POST", doc.Id, "Document"

    ' Update document state
    doc.Status = "Posted"
    doc.IsLocked = False

    ' Register processed operation for idempotency
    opRepo.Add doc.CorrelationId, "POST", doc.Id

    Exit Sub

ERROR_HANDLER:

    ' Rollback all registered operations
    txSvc.Rollback tx
    
    ' Mark document for recovery
    doc.Status = "RecoveryRequired"
    doc.IsLocked = False

    ' Notify error
    MsgBox "POST ERROR: " & ERR.Description

End Sub

' =========================
' CANCEL DOCUMENT
' =========================
' Reverts a previously posted document:
' - Reverses stock movements
' - Creates reversing ledger entry
' - Logs audit action
Public Sub CancelDocument(doc As Ent_Document, lines As Collection)

    On Error GoTo ERROR_HANDLER

    ' Initialize dependencies
    Dim policy As Pol_Document
    Dim stockSvc As Svc_Stock
    Dim audit As Repo_Audit
    Dim repoLedger As Repo_Ledger

    Set policy = New Pol_Document
    Set stockSvc = New Svc_Stock
    Set audit = New Repo_Audit
    Set repoLedger = New Repo_Ledger

    ' Validate document state
    policy.CheckLock doc
    policy.CanCancel doc

    ' Validate input lines
    If lines Is Nothing Or lines.count = 0 Then
        ERR.Raise 500, , "No lines to cancel"
    End If

    ' Lock document for cancellation
    doc.Status = "Cancelling"
    doc.IsLocked = True
    doc.LockedAt = Now

    Dim i As Long
    Dim line As Ent_DocumentLine
    Dim revenue As Double

    ' Process reversal logic
    For i = 1 To lines.count
        
        Set line = lines(i)

        ' Validate line existence
        If line Is Nothing Then
            ERR.Raise 501, , "Line is nothing"
        End If

        ' Reverse stock movements based on document type
        Select Case doc.DocumentType
            
            Case "SALE"
                stockSvc.AddStock line.ProductId, line.Quantity
            
            Case "STOCK_IN"
                stockSvc.RemoveStock line.ProductId, line.Quantity
            
            Case Else
                ERR.Raise 503, , "Unknown document type"
                
        End Select
        
        ' Accumulate revenue for reversal
        revenue = revenue + (line.Quantity * line.Price)
        
    Next i

    ' Create reversing ledger entry for sales
    If doc.DocumentType = "SALE" Then
        
        Dim ledger As Ent_Ledger
        Set ledger = New Ent_Ledger
        
        ledger.Id = GenerateId()
        ledger.CustomerId = 1
        ledger.Amount = -revenue
        ledger.EntryType = "CANCEL"
        ledger.DocumentId = doc.Id
        ledger.CreatedAt = Now
        
        repoLedger.Add ledger
        
    End If

    ' Log cancellation action
    audit.LogAction "CANCEL", doc.Id, "Document"

    ' Update document state
    doc.Status = "Cancelled"
    doc.IsLocked = False

    Exit Sub

ERROR_HANDLER:

    ' Mark document for recovery
    doc.Status = "RecoveryRequired"
    doc.IsLocked = False

    ' Notify error
    MsgBox "CANCEL ERROR: " & ERR.Description

End Sub

' =========================
' HELPERS
' =========================
' Generates incremental unique ID (in-memory)
Private Function GenerateId() As Long
    Static currentId As Long
    currentId = currentId + 1
    GenerateId = currentId
End Function

' Generates correlation ID for idempotency and tracking
Private Function GenerateCorrelationId() As String
    GenerateCorrelationId = Format(Now, "yyyymmddhhmmss") & "_" & Int(Rnd() * 1000)
End Function
