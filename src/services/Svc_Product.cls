VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Svc_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' PRODUCT SERVICE (APPLICATION LAYER)
' =========================================
' Provides access to product-related operations.
' Acts as a thin abstraction over the repository layer.

' Retrieves a product entity by its ID
' Parameters:
' - ProductId: Unique identifier of the product
' Returns:
' - Ent_Product instance if found
' - Nothing if not found
Public Function GetProduct(ProductId As Long) As Ent_Product

    ' Initialize repository dependency
    Dim repo As New Repo_Product
    
    ' Delegate retrieval to repository layer
    Set GetProduct = repo.GetById(ProductId)

End Function
