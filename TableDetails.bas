Attribute VB_Name = "TableDetails"
Option Explicit

' Built on 6/12/2020 3:16:22 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableDetails."

Private Type PrivateType
    Initialized As Boolean
    Dict As Dictionary
    Wkbk As Workbook
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Private Const pColumnHeaderColumn As Long = 1
Private Const pVariableNameColumn As Long = 2
Private Const pVariableTypeColumn As Long = 3
Private Const pKeyColumn As Long = 4
Private Const pFormatColumn As Long = 5
Private Const pHeaderWidth As Long = 5

Private Const pFileName As String = vbNullString
Private Const pWorksheetName As String = vbNullString
Private Const pExternalTableName As String = vbNullString

Public Property Get ColumnHeaderColumn() As Long
    ColumnHeaderColumn = pColumnHeaderColumn
End Property ' ColumnHeaderColumn

Public Property Get VariableNameColumn() As Long
    VariableNameColumn = pVariableNameColumn
End Property ' VariableNameColumn

Public Property Get VariableTypeColumn() As Long
    VariableTypeColumn = pVariableTypeColumn
End Property ' VariableTypeColumn

Public Property Get KeyColumn() As Long
    KeyColumn = pKeyColumn
End Property ' KeyColumn

Public Property Get FormatColumn() As Long
    FormatColumn = pFormatColumn
End Property ' FormatColumn

Public Property Get Headers() As Variant
    Headers = Array( _
        "Column Header", "Variable Name", _
        "Type", "Key", _
        "Format")
End Property ' Headers

Public Property Get Dict() As Dictionary
   Set Dict = This.Dict
End Property ' Dict

Public Property Get SpecificTable() As ListObject
    ' Table in this workbook
    Set SpecificTable = TableDetailsSheet.ListObjects("TableDetailsTable")
End Property ' SpecificTable

Public Property Get Initialized() As Boolean
   Initialized = This.Initialized
End Property ' Initialized

Public Sub Initialize()

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    Dim LocalTable As TableDetails_Table
    Set LocalTable = New TableDetails_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(LocalTable, This.Dict, TableDetails.SpecificTable) Then
        This.Initialized = True
    Else
        ReportError "Error copying TableDetails table", "Routine", RoutineName
        This.Initialized = False
        GoTo Done
    End If

    If Not This.Wkbk Is Nothing Then This.Wkbk.Close
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' TableDetailsInitialize

Public Sub Reset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub ' Reset

Public Property Get HeaderWidth() As Long
    HeaderWidth = pHeaderWidth
End Property ' HeaderWidth

Public Property Get GetVariableNameFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetVariableNameFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetails.Initialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetVariableNameFromColumnHeader = This.Dict(ColumnHeader).VariableName
    Else
        ReportError "Unrecognized ColumnHeader", _
            "Routine", RoutineName, _
            "Column Header", ColumnHeader
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Column Header", ColumnHeader

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetVariableNameFromColumnHeader

Public Property Get GetVariableTypeFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetVariableTypeFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetails.Initialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetVariableTypeFromColumnHeader = This.Dict(ColumnHeader).VariableType
    Else
        ReportError "Unrecognized ColumnHeader", _
            "Routine", RoutineName, _
            "Column Header", ColumnHeader
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Column Header", ColumnHeader

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetVariableTypeFromColumnHeader

Public Property Get GetKeyFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetKeyFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetails.Initialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetKeyFromColumnHeader = This.Dict(ColumnHeader).Key
    Else
        ReportError "Unrecognized ColumnHeader", _
            "Routine", RoutineName, _
            "Column Header", ColumnHeader
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Column Header", ColumnHeader

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetKeyFromColumnHeader

Public Property Get GetFormatFromColumnHeader(ByVal ColumnHeader As String) As String

    Const RoutineName As String = Module_Name & "GetFormatFromColumnHeader"
    On Error GoTo ErrorHandler

    If Not This.Initialized Then TableDetails.Initialize

    If CheckColumnHeaderExists(ColumnHeader) Then
        GetFormatFromColumnHeader = This.Dict(ColumnHeader).Format
    Else
        ReportError "Unrecognized ColumnHeader", _
            "Routine", RoutineName, _
            "Column Header", ColumnHeader
    End If

Done:
    Exit Property
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Column Header", ColumnHeader

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property ' GetFormatFromColumnHeader

Public Function CreateKey(ByVal Record As TableDetails_Table) As String

    Const RoutineName As String = Module_Name & "CreateKey"
    On Error GoTo ErrorHandler

    CreateKey = Record.ColumnHeader

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CreateKey

Public Function TryCopyDictionaryToArray( _
    ByVal Dict As Dictionary, _
    ByRef Ary As Variant _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyDictionaryToArray"
    On Error GoTo ErrorHandler

    TryCopyDictionaryToArray = True

    If Dict.Count = 0 Then
        ReportError "Error copying TableDetails_Table dictionary to array,", "Routine", RoutineName
        TryCopyDictionaryToArray = False
        GoTo Done
    End If

    ReDim Ary(1 To Dict.Count, 1 To 5)

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

Public Function TryCopyArrayToDictionary( _
       ByVal Ary As Variant, _
       ByRef Dict As Dictionary _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record As TableDetails_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Set Record = New TableDetails_Table

            Record.ColumnHeader = Ary(I, pColumnHeaderColumn)
            Record.VariableName = Ary(I, pVariableNameColumn)
            Record.VariableType = Ary(I, pVariableTypeColumn)
            Record.Key = Ary(I, pKeyColumn)
            Record.Format = Ary(I, pFormatColumn)

            Key = TableDetails.CreateKey(Record)

            If Not Dict.Exists(Key) Then
                Dict.Add Key, Record
            Else
                ReportWarning "Duplicate key", "Routine", RoutineName, "Key", Key
                TryCopyArrayToDictionary = False
                GoTo Done
            End If
        Next I

    Else
        ReportError "Invalid Array", "Routine", RoutineName
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

    If Not This.Initialized Then TableDetails.Initialize

    If ColumnHeader = vbNullString Then
        CheckColumnHeaderExists = True
        GoTo Done
    End If

    CheckColumnHeaderExists = This.Dict.Exists(ColumnHeader)

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description ' _
                "Column Header", ColumnHeader

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckColumnHeaderExists

Public Sub FormatArrayAndWorksheet( _
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

