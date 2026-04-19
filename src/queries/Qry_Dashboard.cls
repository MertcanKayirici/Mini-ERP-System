VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Qry_Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' DASHBOARD QUERY (TOTAL STOCK)
' =========================================
' Retrieves total stock value from stock service.
' Includes basic validation and safe fallback on failure.

Public Function GetTotalStock() As Double

    ' Initialize stock service dependency
    Dim svc As New Svc_Stock
    
    On Error GoTo FAIL
    
    ' Retrieve total stock from service layer
    GetTotalStock = svc.GetTotalStock()
    
    ' Validate that total stock is not negative
    If GetTotalStock < 0 Then
        ERR.Raise 3001, "Qry_Dashboard", "Total stock cannot be negative"
    End If
    
    Exit Function

FAIL:
    ' Fallback value in case of error
    GetTotalStock = 0

End Function
