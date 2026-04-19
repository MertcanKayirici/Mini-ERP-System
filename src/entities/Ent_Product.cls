VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ent_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' PRODUCT ENTITY (MASTER DATA MODEL)
' =========================================
' Represents a product within the system.
' This entity is used for pricing, costing,
' and activation state management.

' Primary key - unique identifier for the product
Public Id As Long

' Name of the product
Public Name As String

' Selling price of the product
Public Price As Double

' Cost value used for margin and profitability calculations
Public Cost As Double

' Indicates whether the product is active and available for operations
Public IsActive As Boolean

' Timestamp indicating when the product was created
Public CreatedAt As Date

' =========================
' VALIDATION
' =========================
' Validates product data before persistence or processing
Public Sub Validate()

    ' Ensure product name is provided
    If Name = "" Then
        ERR.Raise vbObjectError + 100, "Product", "Name is required"
    End If

    ' Ensure price is not negative
    If Price < 0 Then
        ERR.Raise vbObjectError + 101, "Product", "Price cannot be negative"
    End If

    ' Ensure cost is not negative
    If Cost < 0 Then
        ERR.Raise vbObjectError + 102, "Product", "Cost cannot be negative"
    End If

End Sub
