VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Qry_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' PRODUCT QUERY (GET PRODUCT NAME)
' =========================================
' Retrieves product name by ProductId using service layer.
' Includes null safety and fallback values for missing or empty data.

Public Function GetProductName(ProductId As Long) As String

    ' Initialize product service
    Dim svc As New Svc_Product
    Dim p As Ent_Product
    
    ' Retrieve product entity from service layer
    Set p = svc.GetProduct(ProductId)
    
    ' Handle case where product does not exist
    If p Is Nothing Then
        GetProductName = "NOT FOUND"
        Exit Function
    End If
    
    ' Ensure safe string return for empty product names
    If Trim(p.Name) = "" Then
        GetProductName = "UNKNOWN"
    Else
        GetProductName = p.Name
    End If

End Function
