Attribute VB_Name = "TableDetails"
Option Explicit

' Built on 12/31/2019 12:09:32 PM
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

Public Property Get TableDetailsDictionary() As Dictionary
   Set TableDetailsDictionary = pTableDetailsDict
End Property

Public Sub TableDetailsReset()
    pInitialized = False
    Set pTableDetailsDict = Nothing
End Sub

Public Function TableDetailsTryCopyTableToDictionary( _
    ByVal Tbl As ListObject, _
    Optional ByRef Dict As Dictionary _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TableDetailsTryCopyTableToDictionary"
    On Error GoTo ErrorHandler
    TableDetailsTryCopyTableToDictionary = True

    Dim Ary As Variant
    Ary = Tbl.DataBodyRange
    If Err.Number <> 0 Then
        MsgBox "The TableDetails table is empty"
        TableDetailsTryCopyTableToDictionary = False
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

    If TableDetailsTryCopyArrayToDictionary(Ary, ThisDict) Then
        ' Success; do nothing
    Else
        ReportError "Error copying array to dictionary", "Routine", RoutineName
        TableDetailsTryCopyTableToDictionary = False
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
End Function ' TableDetailsTryCopyTableToDictionary

Public Function TableDetailsTryCopyDictionaryToTable( _
       ByVal Dict As Dictionary, _
       Optional ByVal Table As ListObject = Nothing, _
       Optional TableCorner As Range = Nothing, _
       Optional TableName As String _
       ) As Boolean

    ' This routine copies a dictionary to a table
    ' If Dict is nothing then use default dictionary
    ' If Table is nothing then build a table using TableCorner and TableName
    ' if Table and TableCorner are both Nothing then use TableDetailsTable

    Const RoutineName As String = Module_Name & "TableDetailsTryCopyDictionaryToTable"
    On Error GoTo ErrorHandler

    TableDetailsTryCopyDictionaryToTable = True

    If Not pInitialized Then TableDetailsInitialize

    Dim ClassName As TableDetails_Table
    Set ClassName = New TableDetails_Table

    '    FormatColumnAsText pFirstColumn, Table, TableCorner

    If Table.TryCopyDictionaryToTable(ClassName, Dict, Table, TableCorner, TableName) Then
        ' Success; do
    Else
        MsgBox "Error copying TableDetails dictionary to table"
        TableDetailsTryCopyDictionaryToTable = False
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TableDetailsTryCopyDictionaryToTable

Private Property Get TableDetailsHeaderWidth() As Long
    TableDetailsHeaderWidth = pHeaderWidth
End Property

Private Property Get TableDetailsHeaders() As Variant
    TableDetailsHeaders = Array("Column Header", "Variable Name", "Formatted?", "Type")
End Property

Private Sub TableDetailsInitialize()

    Const RoutineName As String = Module_Name & "TableDetailsInitialize"
    On Error GoTo ErrorHandler

    pInitialized = True

    Dim TableDetails As TableDetails_Table
    Set TableDetails = New TableDetails_Table

    Set pTableDetailsDict = New Dictionary
    If TableDetailsTryCopyTableToDictionary(TableDetailsTable, pTableDetailsDict) Then
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
End Sub ' TableDetailsInitialize

Private Sub TableDetailsCopyDictionaryToArray( _
    ByVal DetailsDict As Dictionary, _
       ByRef Ary As Variant)

    Const RoutineName As String = Module_Name & "TableDetailsCopyDictionaryToArray"
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
End Sub ' TableDetailsCopyDictionaryToArray

Private Function TableDetailsTryCopyArrayToDictionary( _
       ByVal Ary As Variant, _
       ByRef Dict As Dictionary)

    Const RoutineName As String = Module_Name & "TableDetailsTryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TableDetailsTryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record As TableDetails_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Key = Ary(I, pColumnHeaderColumn)

            If Dict.Exists(Key) Then
                MsgBox "Duplicate key"
                TableDetailsTryCopyArrayToDictionary = False
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

    Else
        Dict.Add Ary, Ary
    End If

    '    Array formatting goes here

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TableDetailsTryCopyArrayToDictionary

