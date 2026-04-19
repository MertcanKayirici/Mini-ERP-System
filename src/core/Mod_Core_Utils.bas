Attribute VB_Name = "Mod_Core_Utils"
Option Explicit

' =========================================
' ID GENERATOR
' =========================================
' Provides a simple in-memory incremental ID generator.
' Uses a static variable to persist state across function calls
' during the Excel session.

Public Function GenerateId() As Long

    ' Static variable retains its value between calls
    Static currentId As Long
    
    ' Increment ID counter
    currentId = currentId + 1
    
    ' Return newly generated ID
    GenerateId = currentId

End Function
