Attribute VB_Name = "Driver"
Option Explicit

Private Const Module_Name As String = "Driver."

Public Sub Main()
    
    Const RoutineName As String = Module_Name & "Main"
    On Error GoTo ErrorHandler
    
    Dim Sheet As Variant
    
    For Each Sheet In ThisWorkbook.Worksheets
        ClassBuilder.ClassBuilder _
            Sheet.ListObjects(1), _
            Sheet.ListObjects(2)
        ModuleBuilder.ModuleBuilder _
            Sheet.ListObjects(1), _
            Sheet.ListObjects(2)
    Next Sheet

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    CloseErrorFile
End Sub ' Main
