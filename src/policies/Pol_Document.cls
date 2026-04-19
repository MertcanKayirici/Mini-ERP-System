VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pol_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' DOCUMENT POLICY MODULE
' =========================================
' Contains business rules and validation logic
' for document operations such as posting, cancellation,
' locking, and stock/product validation.

' =========================
' CAN POST
' =========================
' Validates whether a document can be posted
Public Function CanPost(doc As Ent_Document) As Boolean

    ' Ensure document is in Draft state
    If doc.Status <> "Draft" Then
        ERR.Raise vbObjectError + 100, "Pol_Document", "Document is not in Draft state"
    End If

    ' Prevent posting if document is locked
    If doc.IsLocked = True Then
        ERR.Raise vbObjectError + 101, "Pol_Document", "Document is locked"
    End If

    CanPost = True

End Function


' =========================
' CAN CANCEL
' =========================
' Validates whether a document can be cancelled
Public Function CanCancel(doc As Ent_Document) As Boolean

    ' Only posted documents can be cancelled
    If doc.Status <> "Posted" Then
        ERR.Raise vbObjectError + 102, "Pol_Document", "Only posted documents can be cancelled"
    End If

    ' Prevent cancellation if document is locked
    If doc.IsLocked = True Then
        ERR.Raise vbObjectError + 103, "Pol_Document", "Document is locked"
    End If

    CanCancel = True

End Function


' =========================
' CHECK LOCK (FINAL + TTL)
' =========================
' Validates lock state and applies TTL-based auto-unlock logic
Public Sub CheckLock(doc As Ent_Document)

    ' Exit if document is not locked
    If doc.IsLocked = False Then Exit Sub

    ' TTL (Time-To-Live): 5 minutes
    ' If lock duration exceeds TTL, automatically unlock
    If Now - doc.LockedAt > TimeSerial(0, 5, 0) Then
        
        ' Force unlock
        doc.IsLocked = False
        
        ' Log forced unlock action
        Dim audit As Repo_Audit
        Set audit = New Repo_Audit
        
        audit.LogAction "FORCE_UNLOCK", doc.Id, "Document"
        
    Else
        ' Prevent operation if still within lock duration
        ERR.Raise vbObjectError + 300, "Pol_Document", "Document is locked"
    End If

End Sub


' =========================
' CHECK PRODUCT ACTIVE
' =========================
' Ensures that the product exists and is active
Public Sub CheckProductActive(ProductId As Long)

    Dim repo As Repo_Product
    Set repo = New Repo_Product
    
    Dim p As Ent_Product
    Set p = repo.GetById(ProductId)
    
    ' Validate product existence
    If p Is Nothing Then
        ERR.Raise vbObjectError + 301, "Pol_Document", "Product not found"
    End If
    
    ' Validate product is active
    If p.IsActive = False Then
        ERR.Raise vbObjectError + 300, "Pol_Document", "Product is inactive"
    End If

End Sub


' =========================
' CHECK STOCK AVAILABILITY
' =========================
' Ensures sufficient stock is available for the operation
Public Sub CheckStockAvailability(ProductId As Long, Quantity As Double)

    Dim stockSvc As Svc_Stock
    Set stockSvc = New Svc_Stock
    
    ' Retrieve current stock level
    Dim currentStock As Double
    currentStock = stockSvc.GetCurrentStock(ProductId)

    ' Validate stock sufficiency
    If currentStock < Quantity Then
        ERR.Raise vbObjectError + 200, "Pol_Document", "Insufficient stock"
    End If

End Sub
