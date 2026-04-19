VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Svc_Governance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================
' GOVERNANCE MODULE (ARCHITECTURE VALIDATION)
' =========================================
' Enforces architectural rules such as:
' - Layer dependency restrictions
' - Naming conventions
' - Query layer read-only enforcement

Public Sub RunFullGovernance()

    On Error GoTo ERR_HANDLER

    ' Execute all governance checks
    ValidateDependencies
    ValidateNaming
    ValidateQueryLayer

    ' Notify success if all checks pass
    MsgBox "Governance PASSED", vbInformation
    Exit Sub

ERR_HANDLER:
    ' Notify failure with error details
    MsgBox "Governance FAILED: " & ERR.Description, vbCritical

End Sub


' =========================
' DEPENDENCY
' =========================
' Validates forbidden dependencies between layers
Private Sub ValidateDependencies()

    CheckForbidden "Repo_", "Svc_"
    CheckForbidden "Repo_", "Pol_"
    CheckForbidden "Qry_", "Repo_"
    CheckForbidden "Qry_", "Svc_"

End Sub


' =========================
' NAMING
' =========================
' Ensures all modules follow required naming conventions
Private Sub ValidateNaming()

    Dim comp As Object
    
    ' Iterate through all VBA components
    For Each comp In ThisWorkbook.VBProject.VBComponents
        
        ' Validate prefix compliance
        If Not IsValidPrefix(comp.Name) Then
            ERR.Raise 1001, "Governance", "Invalid naming: " & comp.Name
        End If
        
    Next comp

End Sub


' Validates whether a component name starts with an allowed prefix
Private Function IsValidPrefix(Name As String) As Boolean

    Dim prefixes As Variant
    prefixes = Array("Svc_", "Repo_", "Pol_", "Ent_", "Qry_", "frm_", "Const_", "Infra_")
    
    Dim i As Integer
    
    ' Check against allowed prefixes
    For i = LBound(prefixes) To UBound(prefixes)
        If Left(Name, Len(prefixes(i))) = prefixes(i) Then
            IsValidPrefix = True
            Exit Function
        End If
    Next i

    IsValidPrefix = False

End Function


' =========================
' QUERY RULES
' =========================
' Ensures query layer does not perform write operations
Private Sub ValidateQueryLayer()

    Dim comp As Object
    Dim code As String
    
    ' Iterate through all VBA components
    For Each comp In ThisWorkbook.VBProject.VBComponents
        
        ' Target only query modules
        If Left(comp.Name, 4) = "Qry_" Then
            
            ' Read module code
            If comp.CodeModule.CountOfLines > 0 Then
                code = comp.CodeModule.lines(1, comp.CodeModule.CountOfLines)
            End If
            
            ' Detect forbidden write operations
            If InStr(1, code, "Add(", vbTextCompare) > 0 _
            Or InStr(1, code, "Update(", vbTextCompare) > 0 _
            Or InStr(1, code, "Delete(", vbTextCompare) > 0 Then
                
                ERR.Raise 1100, "Governance", _
                "Query layer writes data: " & comp.Name
                
            End If
            
        End If
        
    Next comp

End Sub


' =========================
' CORE CHECK
' =========================
' Validates forbidden dependencies between modules
Private Sub CheckForbidden(prefix1 As String, prefix2 As String)

    Dim comp As Object
    Dim code As String
    
    ' Iterate through all VBA components
    For Each comp In ThisWorkbook.VBProject.VBComponents
        
        ' Check components with specified prefix
        If InStr(1, comp.Name, prefix1, vbTextCompare) > 0 Then
            
            ' Read component code
            If comp.CodeModule.CountOfLines > 0 Then
                code = comp.CodeModule.lines(1, comp.CodeModule.CountOfLines)
                
                ' Detect forbidden dependency usage
                If InStr(1, code, prefix2, vbTextCompare) > 0 Then
                    ERR.Raise vbObjectError + 900, "Governance", _
                    "Forbidden dependency: " & comp.Name & " -> " & prefix2
                End If
                
            End If
            
        End If
        
    Next comp

End Sub


' =========================
' GOVERNANCE REPORT
' =========================
' Executes governance checks and writes results into a worksheet report
Public Sub RunGovernanceWithReport()

    Dim ws As Worksheet
    
    ' ?????? ?????? ??? ?????? ?? ??????? ??? ?? ??? ??????
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("GovernanceReport")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "GovernanceReport"
    End If
    
    ' Clear previous report data
    ws.Cells.Clear
    
    ' Set report headers
    ws.Cells(1, 1).value = "Check"
    ws.Cells(1, 2).value = "Result"
    
    Dim row As Long
    row = 2
    
    ' Dependency check
    If RunSafeCheck("ValidateDependencies") Then
        WriteResult ws, row, "Dependencies", "PASS"
    Else
        WriteResult ws, row, "Dependencies", "FAIL"
    End If
    
    ' Naming check
    If RunSafeCheck("ValidateNaming") Then
        WriteResult ws, row, "Naming", "PASS"
    Else
        WriteResult ws, row, "Naming", "FAIL"
    End If
    
    ' Query layer check
    If RunSafeCheck("ValidateQueryLayer") Then
        WriteResult ws, row, "Query Layer", "PASS"
    Else
        WriteResult ws, row, "Query Layer", "FAIL"
    End If
    
    ' Notify report generation
    MsgBox "Governance Report is ready"

End Sub


' Executes a validation method safely and returns success/failure
Private Function RunSafeCheck(methodName As String) As Boolean

    On Error GoTo FAIL
    
    ' Dynamically invoke validation method
    CallByName Me, methodName, VbMethod
    
    RunSafeCheck = True
    Exit Function

FAIL:
    RunSafeCheck = False

End Function


' Writes a single result row into the governance report
Private Sub WriteResult(ws As Worksheet, ByRef row As Long, testName As String, result As String)

    ws.Cells(row, 1).value = testName
    ws.Cells(row, 2).value = result
    
    ' Apply color coding based on result
    If result = "PASS" Then
        ws.Cells(row, 2).Interior.color = vbGreen
    Else
        ws.Cells(row, 2).Interior.color = vbRed
    End If
    
    row = row + 1

End Sub

