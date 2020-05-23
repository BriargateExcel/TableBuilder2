Attribute VB_Name = "Driver"
Option Explicit
'@Folder "Builder"
Private Const Module_Name As String = "Driver."

Public Sub BuildModules()
    
    Const RoutineName As String = Module_Name & "Main"
    On Error GoTo ErrorHandler
    
    Dim Sheet As Variant
    
    Dim BasicDict As Dictionary
    Dim BasicsTable As ListObject
    Dim TableBasics As TableBasics_Table
    Set TableBasics = New TableBasics_Table
    
    Dim DetailsDict As Dictionary
    Dim DetailsTable As ListObject
    Dim TableDetails As TableDetails_Table
    Set TableDetails = New TableDetails_Table
    
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name <> "VBA Make File" Then
            Set BasicsTable = Sheet.ListObjects(2)
            If BasicsTable.HeaderRowRange(1, 1) <> "Table Name" Then
                Set BasicsTable = Sheet.ListObjects(1)
                Set DetailsTable = Sheet.ListObjects(2)
            Else
                Set BasicsTable = Sheet.ListObjects(2)
                Set DetailsTable = Sheet.ListObjects(1)
            End If
            
            If Table.TryCopyTableToDictionary(TableBasics, BasicDict, BasicsTable) Then
                ' Success; do nothing
            Else
                ReportError "Error copying TableBasics to dictionary", "Routine", RoutineName
            End If
            
            If Table.TryCopyTableToDictionary(TableDetails, DetailsDict, DetailsTable) Then
                ' Success; do nothing
            Else
                ReportError "Error copying Table to dictionary", "Routine", RoutineName
            End If
            
            ClassBuilder.ClassBuilder DetailsDict, BasicDict
            
            ModuleBuilder.ModuleBuilder DetailsDict, BasicDict
        End If
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
