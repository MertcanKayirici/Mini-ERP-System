VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ent_StockMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' STOCK MOVEMENT ENTITY (INVENTORY MODEL)
' =========================================
' Represents a stock transaction (inbound or outbound).
' This entity is used to track inventory changes and
' their relationship to business documents.

' Primary key - unique identifier for the stock movement
Public Id As Long

' Foreign key reference to the related product
Public ProductId As Long

' Quantity of the product moved in this transaction
Public Quantity As Double

' Defines the type of movement (e.g., IN for incoming, OUT for outgoing)
Public MovementType As String

' Reference to the associated document (e.g., Invoice, Order)
Public DocumentId As Long

' Timestamp indicating when the stock movement was created
Public CreatedAt As Date
