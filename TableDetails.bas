Attribute VB_Name = "TableDetails"
Option Explicit

' Built on 3/8/2020 9:42:44 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableDetails."

Private Type TableDetailsType
    Initialized As Boolean
    Dict As Dictionary
End Type

Private This As TableDetailsType

' No application specific declarations found

Private Const pColumnHeaderColumn As Long = 1
Private Const pVariableNameColumn As Long = 2
Private Const pVariableTypeColumn As Long = 3
Private Const pKeyColumn As Long = 4
Private Const pFormatColumn As Long = 5
Private Const pHeaderWidth As Long = 5

Public Property Get TableDetailsColumnHeaderColumn() As Long
    TableDetailsColumnHeaderColumn = pColumnHeaderColumn
End Property

Public Property Get TableDetailsVariableNameColumn() As Long
    TableDetailsVariableNameColumn = pVariableNameColumn
End Property

Public Property Get TableDetailsVariableTypeColumn() As Long
    TableDetailsVariableTypeColumn = pVariableTypeColumn
End Property

Public Property Get TableDetailsKeyColumn() As Long
    TableDetailsKeyColumn = pKeyColumn
End Property

Public Property Get TableDetailsFormatColumn() As Long
    TableDetailsFormatColumn = pFormatColumn
End Property

Public Property Get TableDetailsHeaders() As Variant
    TableDetailsHeaders = Array( _
        "Column Header", "Variable Name", _
        "Type", "Key", _
        "Format")
End Property

Public Property Get GetVariableNameFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetVariableNameFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetailsInitialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetVariableNameFromColumnHeader = This.Dict(ColumnHeader).VariableName
    Else
        ReportError "Unrecognized ColumnHeader", "Routine", RoutineName
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetVariableNameFromColumnHeader

Public Property Get GetVariableTypeFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetVariableTypeFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetailsInitialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetVariableTypeFromColumnHeader = This.Dict(ColumnHeader).VariableType
    Else
        ReportError "Unrecognized ColumnHeader", "Routine", RoutineName
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetVariableTypeFromColumnHeader

Public Property Get GetKeyFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetKeyFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetailsInitialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetKeyFromColumnHeader = This.Dict(ColumnHeader).Key
    Else
        ReportError "Unrecognized ColumnHeader", "Routine", RoutineName
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetKeyFromColumnHeader

Public Property Get GetFormatFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetFormatFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetailsInitialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetFormatFromColumnHeader = This.Dict(ColumnHeader).Format
    Else
        ReportError "Unrecognized ColumnHeader", "Routine", RoutineName
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetFormatFromColumnHeader

Public Property Get TableDetailsDictionary() As Dictionary
   Set TableDetailsDictionary = This.Dict
End Property

Public Property Get TableDetailsTable() As ListObject

    ' Change the table reference if the table is in another workbook

    Set TableDetailsTable = TableDetailsSheet.ListObjects("TableDetailsTable")
End Property

Public Property Get TableDetailsInitialized() As Boolean
   TableDetailsInitialized = This.Initialized
End Property

Public Sub TableDetailsInitialize()

    Const RoutineName As String = Module_Name & "TableDetailsInitialize"
    On Error GoTo ErrorHandler
    Dim TableDetails As TableDetails_Table
    Set TableDetails = New TableDetails_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(TableDetails, TableDetailsTable, This.Dict) Then
        This.Initialized = True
    Else
        ReportError "Error copying TableDetails table", "Routine", RoutineName
        This.Initialized = False
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

Public Sub TableDetailsReset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub

Public Property Get TableDetailsHeaderWidth() As Long
    TableDetailsHeaderWidth = pHeaderWidth
End Property

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
        Ary(I, pKeyColumn) = Record.Key
        Ary(I, pFormatColumn) = Record.Format

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
       ByRef Dict As Dictionary _
       ) As Boolean

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
                ReportWarning "Duplicate key", "Routine", RoutineName, "Key", Key
                TableDetailsTryCopyArrayToDictionary = False
                GoTo Done
            Else
                Set Record = New TableDetails_Table

                Record.ColumnHeader = Ary(I, pColumnHeaderColumn)
                Record.VariableName = Ary(I, pVariableNameColumn)
                Record.VariableType = Ary(I, pVariableTypeColumn)
                Record.Key = Ary(I, pKeyColumn)
                Record.Format = Ary(I, pFormatColumn)

                Dict.Add Key, Record
            End If
        Next I

    Else
        Dict.Add Ary, Ary
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TableDetailsTryCopyArrayToDictionary

Public Function CheckColumnHeaderExists(ByVal ColumnHeader As String) As Boolean _

    Const RoutineName As String = Module_Name & "CheckColumnHeaderExists"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetailsInitialize

    If ColumnHeader = vbNullString Then
        CheckColumnHeaderExists = True
        Exit Function
    End If

    CheckColumnHeaderExists = This.Dict.Exists(ColumnHeader)

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckColumnHeaderExists

Public Sub TableDetailsFormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "TableDetailsFormatArrayAndWorksheet"
    On Error GoTo ErrorHandler


Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' TableDetailsFormatArrayAndWorksheet

' No application unique routines found

