VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ent_Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' CUSTOMER LEDGER ENTRY (ACCOUNTING MODEL)
' =========================================
' Represents a financial transaction associated with a customer.
' This entity is used to track balances, debit/credit movements,
' and relationships with business documents.

' Primary key - unique identifier for the ledger entry
Public Id As Long

' Foreign key reference to the related customer
Public CustomerId As Long

' Monetary value of the transaction
Public Amount As Double

' Defines the type of ledger entry (e.g., Debit, Credit)
Public EntryType As String

' Reference to the associated document (e.g., Invoice, Payment)
Public DocumentId As Long

' Timestamp indicating when the ledger entry was created
Public CreatedAt As Date
