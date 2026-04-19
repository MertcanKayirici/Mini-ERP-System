VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ent_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' DOCUMENT ENTITY (DATA MODEL)
' =========================================
' Represents a business document within the system.
' This entity is used to manage document lifecycle,
' locking state, and integration consistency.

' Primary key - unique identifier for the document
Public Id As Long

' Defines the type of document (e.g., Invoice, Order, Receipt)
Public DocumentType As String

' Represents the current lifecycle state of the document
' (e.g., Draft, Approved, Cancelled)
Public Status As String

' Correlation identifier used for cross-system tracking,
' logging, and idempotent operations
Public CorrelationId As String

' Indicates whether the document is currently locked for editing
Public IsLocked As Boolean

' Stores the user or system identifier that locked the document
Public LockedBy As String

' Timestamp of when the document was locked
Public LockedAt As Date

' Flag indicating whether stock movements have been generated
' for this document
Public StockCreated As Boolean

' Flag indicating whether audit log entries have been created
' for this document
Public AuditCreated As Boolean

' Flag indicating whether accounting (ledger) entries have been created
Public LedgerCreated As Boolean
