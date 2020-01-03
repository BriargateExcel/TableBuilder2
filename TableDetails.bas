Attribute VB_Name = "TableDetails"
Option Explicit

' Built on 1/1/2020 9:32:28 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableDetails."

Private pInitialized As Boolean
Private pTableDetailsDict As Dictionary

Private Const pColumnHeaderColumn As Long = 1
Private Const pVariableNameColumn As Long = 2
Private Const pVariableTypeColumn As Long = 3
Private Const pHeaderWidth As Long = 3

Public Property Get TableDetailsTable() As ListObject
    Set TableDetailsTable = TableDetailsSheet.ListObjects("TableDetailsTable")
End Property

Public Property Get TableDetailsDictionary() As Dictionary
   Set TableDetailsDictionary = pTableDetailsDict
End Property

Public Property Get TableDetailsInitialized() As Boolean
   TableDetailsInitialized = pInitialized
End Property

Public Sub TableDetailsReset()
    pInitialized = False
    Set pTableDetailsDict = Nothing
End Sub

Public Property Get TableDetailsHeaderWidth() As Long
    TableDetailsHeaderWidth = pHeaderWidth
End Property

Public Property Get TableDetailsHeaders() As Variant
    TableDetailsHeaders = Array("Column Header", "Variable Name", "Type")
End Property

Public Sub TableDetailsInitialize()

    Const RoutineName As String = Module_Name & "TableDetailsInitialize"
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
End Sub ' TableDetailsInitialize

Public Function TableDetailsTryCopyDictionaryToArray( _
    ByVal Dict As Dictionary, _
    ByRef Ary As Variant _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TableDetailsTryCopyDictionaryToArray"
    On Error GoTo ErrorHandler

    TableDetailsTryCopyDictionaryToArray = True

    If Dict.Count = 0 Then
        ReportError "Error copying TableDetails dictionary to array,", "Routine", RoutineName
        TableDetailsTryCopyDictionaryToArray = False
        GoTo Done
    End If
    Dim I As Long
    I = 1

    Dim Record As TableDetails_Table
    Dim Entry As Variant
    For Each Entry In Dict.Keys
        Set Record = Dict.Item(Entry)

        Ary(I, pColumnHeaderColumn) = Record.ColumnHeader
        Ary(I, pVariableNameColumn) = Record.VariableName
        Ary(I, pVariableTypeColumn) = Record.VariableType

        I = I + 1
    Next Entry
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TableDetailsTryCopyDictionaryToArray

Public Function TableDetailsTryCopyArrayToDictionary( _
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

Public Sub TableDetailsFormatWorksheet(ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "TableDetailsFormatWorksheet"
    On Error GoTo ErrorHandler

    ' Worksheet formatting goes here

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' TableDetailsFormatWorksheet

''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  '
'             End of Generated code                '
'            Start unique code here                '
'                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''

