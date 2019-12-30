Attribute VB_Name = "TableDetails"
'Attribute VB_Name = "TableDetails"
Option Explicit

' Built on 12/30/2019 7:27:22 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableDetails."

Private pInitialized As Boolean
Private pTableDetailsDict As Dictionary

Private Const pColumnHeaderColumn As Long = 1
Private Const pVariableNameColumn As Long = 2
Private Const pFormattedColumn As Long = 3
Private Const pVariableTypeColumn As Long = 4
Private Const pHeaderWidth As Long = 4

Public Property Get TableDetailsTable() As ListObject
    Set TableDetailsTable = TableDetailsSheet.ListObjects("TableDetailsTable")
End Property

Public Sub ResetTableDetails()
    pInitialized = False
    Set pTableDetailsDict = Nothing
End Sub

Public Function TryCopyTableToDictionary( _
    ByVal Tbl As ListObject, _
    Optional ByRef Dict As Dictionary _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyTableToDictionary"
    On Error GoTo ErrorHandler
    TryCopyTableToDictionary = True

    Dim Ary As Variant
    Ary = Tbl.DataBodyRange
    If Err.Number <> 0 Then
        MsgBox "The TableDetails table is empty"
        TryCopyTableToDictionary = False
        GoTo Done
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    Dim ThisDict As Dictionary
    If Dict Is Nothing Then
        Set ThisDict = New Dictionary
    Else
        Set ThisDict = pTableDetailsDict
    End If

    If TableDetails.TryCopyArrayToDictionary(Ary, ThisDict) Then
        ' Success; do nothing
    Else
        ReportError "Error copying array to dictionary", "Routine", RoutineName
        TryCopyTableToDictionary = False
        GoTo Done
    End If

    Set Dict = ThisDict

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyTableToDictionary

Public Property Get Headers() As Variant
    Headers = Array("ColumnHeader", "VariableName", "Formatted", "VariableType")
End Property

Public Property Get TableDetailsDictionary() As Dictionary
   Set TableDetailsDictionary = pTableDetailsDict
End Property

Public Property Get TableDetailsHeaderWidth() As Long
    TableDetailsHeaderWidth = pHeaderWidth
End Property

Public Property Get TableDetailsInitialized() As Boolean
    TableDetailsInitialized = pInitialized
End Property

Public Sub Initialize()

    ' This routine loads the dictionary

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    pInitialized = True

    Dim TableDetails As TableDetails_Table
    Set TableDetails = New TableDetails_Table
    Set pTableDetailsDict = New Dictionary
    If Table.TryCopyTableToDictionary(TableDetails, TableDetailsTable, pTableDetailsDict) Then
        ' Success; do nothing
    Else
        MsgBox "Error copying TableDetails table"
        pInitialized = False
        GoTo Done
    End If

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' Initialize

Public Sub CopyDictionaryToArray( _
    ByVal DetailsDict As Dictionary, _
       ByRef Ary As Variant)

    ' loads TableDetails Dict into Ary

    Const RoutineName As String = Module_Name & "CopyDictionaryToArray"
    On Error GoTo ErrorHandler

    Dim I As Long
    I = 1

    Dim Record As TableDetails_Table
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        Set Record = DetailsDict.Item(Entry)

        Ary(I, pColumnHeaderColumn) = Record.ColumnHeader
        Ary(I, pVariableNameColumn) = Record.VariableName
        Ary(I, pFormattedColumn) = Record.Formatted
        Ary(I, pVariableTypeColumn) = Record.VariableType

        I = I + 1
    Next Entry
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' CopyDictionaryToArray

Public Function TryCopyArrayToDictionary( _
       ByVal Ary As Variant, _
       ByRef Dict As Dictionary)

    ' Copy TableDetails array to dictionary

    Const RoutineName As String = Module_Name & "TryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record As TableDetails_Table

    For I = 1 To UBound(Ary, 1)
        Key = Ary(I, pColumnHeaderColumn)

        If Dict.Exists(Key) Then
            MsgBox "Duplicate key"
            TryCopyArrayToDictionary = False
            GoTo Done
        Else
            Set Record = New TableDetails_Table

            Record.ColumnHeader = Ary(I, pColumnHeaderColumn)
            Record.VariableName = Ary(I, pVariableNameColumn)
            Record.Formatted = IIf(Ary(I, pFormattedColumn) = "Yes", True, False)
            Record.VariableType = Ary(I, pVariableTypeColumn)

            Dict.Add Key, Record
        End If
    Next I

    '    Array formatting goes here

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyArrayToDictionary

Public Function TryCopyDictionaryToTable( _
       ByVal Dict As Dictionary, _
       Optional ByVal Table As ListObject = Nothing, _
       Optional TableCorner As Range = Nothing, _
       Optional TableName As String _
       ) As Boolean

    ' This routine copies a dictionary to a table
    ' If Dict is nothing then use default dictionary
    ' If Table is nothing then build a table using TableCorner and TableName
    ' if Table and TableCorner are both Nothing then use TableDetailsTable

    Const RoutineName As String = Module_Name & "TryCopyDictionaryToTable"
    On Error GoTo ErrorHandler

    TryCopyDictionaryToTable = True

    If Not pInitialized Then TableDetails.Initialize

    Dim ClassName As TableDetails_Table
    Set ClassName = New TableDetails_Table

    '    FormatColumnAsText pFirstColumn, Table, TableCorner

    If Table.TryCopyDictionaryToTable(ClassName, Dict, Table, TableCorner, TableName) Then
        ' Success; do
    Else
        MsgBox "Error copying TableDetails dictionary to table"
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyDictionaryToTable



