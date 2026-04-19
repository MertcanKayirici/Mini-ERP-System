VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ent_DocumentLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' DOCUMENT LINE ENTITY (DETAIL MODEL)
' =========================================
' Represents a single line item within a document.
' This entity is responsible for product-level data such as
' quantity, pricing, cost calculations, and stock integration.

' Primary key - unique identifier for the document line
Public Id As Long

' Foreign key reference to the parent document
Public DocumentId As Long

' Identifier of the associated product
Public ProductId As Long

' Quantity of the product in this line item
Public Quantity As Double

' Unit price assigned to the product
Public Price As Double

' Cost value used for margin and profitability calculations
Public Cost As Double

' Flag indicating whether stock movement has been created
' for this specific line item
Public StockCreated As Boolean

' Flag indicating whether audit log entries have been created
' for this specific line item
Public AuditCreated As Boolean
