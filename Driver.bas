Attribute VB_Name = "Driver"
Option Explicit

Private Const Module_Name As String = "Driver."

Public Sub BuildModules()
    
    Const RoutineName As String = Module_Name & "Main"
    On Error GoTo ErrorHandler
    
    Dim BasicDict As Dictionary
    Dim DetailsDict As Dictionary
    Dim TableName As String
    Dim ClassName As String
    Dim Sheet As Variant
    Dim DetailsTable As ListObject
    Dim BasicsTable As ListObject
    Dim TableDetails As TableDetails_Table
    Set TableDetails = New TableDetails_Table
    
    Dim TableBasics As TableBasics_Table
    Set TableBasics = New TableBasics_Table
    
    For Each Sheet In ThisWorkbook.Worksheets
        Set BasicsTable = Sheet.ListObjects(2)
        If BasicsTable.HeaderRowRange(1, 1) <> "Table Name" Then
            Set BasicsTable = Sheet.ListObjects(1)
            Set DetailsTable = Sheet.ListObjects(2)
        Else
            Set BasicsTable = Sheet.ListObjects(2)
            Set DetailsTable = Sheet.ListObjects(1)
        End If
        
        If Table.TryCopyTableToDictionary(TableBasics, BasicsTable, BasicDict) Then
            ' Success; do nothing
        Else
            ReportError "Error copying TableBasics to dictionary", "Routine", RoutineName
        End If
        
        TableName = BasicDict.Items(0).TableName
        ClassName = TableName & "_Table"
        
        If Table.TryCopyTableToDictionary(TableDetails, DetailsTable, DetailsDict) Then
            ' Success; do nothing
        Else
            ReportError "Error copying Table to dictionary", "Routine", RoutineName
        End If
        
        ClassBuilder.ClassBuilder DetailsDict, TableName, ClassName
        
        ModuleBuilder.ModuleBuilder DetailsDict, TableName, ClassName
    Next Sheet

    MsgBox "Files built", vbOKOnly
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    CloseErrorFile
End Sub ' Main
