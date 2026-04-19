VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repo_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' PRODUCT REPOSITORY (DATA ACCESS LAYER)
' =========================================
' Responsible for managing product persistence,
' retrieval, and state changes in the Products worksheet.
' Encapsulates CRUD-like operations for product entities.

' =========================
' ADD PRODUCT
' =========================
' Creates a new product record
' - Validates product before saving
' - Generates unique ID
' - Persists product data into worksheet
Public Sub Add(p As Ent_Product)

    ' Validate entity before persistence
    p.Validate
    
    ' Assign unique ID
    p.Id = GenerateId()

    Dim ws As Worksheet
    
    ' Reference Products worksheet
    Set ws = ThisWorkbook.Sheets("Products")

    Dim r As Long
    
    ' Determine next available row
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1

    ' Persist product fields
    ws.Cells(r, 1).value = p.Id
    ws.Cells(r, 2).value = p.Name
    ws.Cells(r, 3).value = p.Price
    ws.Cells(r, 4).value = p.Cost
    ws.Cells(r, 5).value = p.IsActive
    ws.Cells(r, 6).value = p.CreatedAt

End Sub


' =========================
' UPDATE PRODUCT
' =========================
' Updates an existing product record
' - Validates product before update
' - Searches by Product ID
' - Throws error if product is not found
Public Sub Update(p As Ent_Product)

    ' Validate entity before update
    p.Validate

    Dim ws As Worksheet
    
    ' Reference Products worksheet
    Set ws = ThisWorkbook.Sheets("Products")

    Dim i As Long

    ' Iterate through existing records
    For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        
        ' Match record by Product ID
        If ws.Cells(i, 1).value = p.Id Then
            
            ' Update product fields
            ws.Cells(i, 2).value = p.Name
            ws.Cells(i, 3).value = p.Price
            ws.Cells(i, 4).value = p.Cost
            ws.Cells(i, 5).value = p.IsActive
            
            Exit Sub
            
        End If
        
    Next i

    ' Raise error if product is not found
    ERR.Raise vbObjectError + 200, "Repo_Product", "Product not found"

End Sub


' =========================
' DEACTIVATE PRODUCT
' =========================
' Performs soft delete by setting IsActive = False
Public Sub Deactivate(ProductId As Long)

    Dim ws As Worksheet
    
    ' Reference Products worksheet
    Set ws = ThisWorkbook.Sheets("Products")

    Dim i As Long

    ' Iterate through records to find target product
    For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        
        ' Match by Product ID
        If ws.Cells(i, 1).value = ProductId Then
            
            ' Mark product as inactive (soft delete)
            ws.Cells(i, 5).value = False
            
            Exit Sub
            
        End If
        
    Next i

End Sub


' =========================
' GET PRODUCT BY ID
' =========================
' Retrieves a product entity by its ID
' Returns:
' - Ent_Product instance if found
' - Nothing if not found
Public Function GetById(ProductId As Long) As Ent_Product

    Dim ws As Worksheet
    
    ' Reference Products worksheet
    Set ws = ThisWorkbook.Sheets("Products")
    
    Dim lastRow As Long
    
    ' Determine last used row
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim i As Long
    
    ' Iterate through records
    For i = 2 To lastRow
        
        ' Match by Product ID
        If CLng(ws.Cells(i, 1).value) = ProductId Then
            
            Dim p As New Ent_Product
            
            ' Map worksheet data to entity fields
            p.Id = ws.Cells(i, 1).value
            p.Name = ws.Cells(i, 2).value
            p.Price = ws.Cells(i, 3).value
            p.Cost = ws.Cells(i, 4).value
            
            ' Ensure safe boolean conversion
            p.IsActive = (ws.Cells(i, 5).value = True)
            
            Set GetById = p
            Exit Function
            
        End If
        
    Next i

    ' Return Nothing if product is not found
    Set GetById = Nothing

End Function


' =========================
' GENERATE ID
' =========================
' Generates incremental unique ID (in-memory)
' NOTE:
' - Resets when application restarts
' - Suitable for simple/local scenarios
Private Function GenerateId() As Long

    ' Static variable preserves state during session
    Static currentId As Long
    
    ' Increment ID counter
    currentId = currentId + 1
    
    ' Return generated ID
    GenerateId = currentId

End Function
