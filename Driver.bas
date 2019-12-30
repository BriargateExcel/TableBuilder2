Attribute VB_Name = "Driver"
Option Explicit

Private Const Module_Name As String = "Driver."

Public Sub Main()
    
    Const RoutineName As String = Module_Name & "Main"
    On Error GoTo ErrorHandler
    
    Dim BasicDict As Dictionary
    Dim DetailsDict As Dictionary
    Dim TableName As String
    Dim ClassName As String
    Dim Sheet As Variant
    
    For Each Sheet In ThisWorkbook.Worksheets
        If TableBasics.TryCopyTableToDictionary(Sheet.ListObjects(2), BasicDict) Then
            ' Success; do nothing
        Else
            ReportError "Error copying TableBasics to dictionary", "Routine", RoutineName
        End If
        
        TableName = BasicDict.Items(0)
        ClassName = TableName & "_Table"
        
        If TableDetails.TryCopyTableToDictionary(Sheet.ListObjects(1), DetailsDict) Then
            ' Success; do nothing
        Else
            ReportError "Error copying Table to dictionary", "Routine", RoutineName
        End If
        
        ClassBuilder.ClassBuilder DetailsDict, TableName, ClassName
        
        ModuleBuilder.ModuleBuilder DetailsDict, TableName, ClassName
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
