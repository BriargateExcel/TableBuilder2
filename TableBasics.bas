Attribute VB_Name = "TableBasics"
Option Explicit

' Built on 4/9/2020 4:36:21 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableBasics."

Private Type PrivateType
    Initialized As Boolean
    Dict As Dictionary
    Wkbk As Workbook
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Private Const pTableNameColumn As Long = 1
Private Const pFileNameColumn As Long = 2
Private Const pWorksheetNameColumn As Long = 3
Private Const pExternalTableNameColumn As Long = 4
Private Const pSkipColumn As Long = 5
Private Const pHeaderWidth As Long = 5

Private Const pFileName As String = vbNullString
Private Const pWorksheetName As String = vbNullString
Private Const pExternalTableName As String = vbNullString

Public Property Get TableNameColumn() As Long
    TableNameColumn = pTableNameColumn
End Property ' TableNameColumn

Public Property Get FileNameColumn() As Long
    FileNameColumn = pFileNameColumn
End Property ' FileNameColumn

Public Property Get WorksheetNameColumn() As Long
    WorksheetNameColumn = pWorksheetNameColumn
End Property ' WorksheetNameColumn

Public Property Get ExternalTableNameColumn() As Long
    ExternalTableNameColumn = pExternalTableNameColumn
End Property ' ExternalTableNameColumn

Public Property Get SkipColumn() As Long
    SkipColumn = pSkipColumn
End Property ' SkipColumn

Public Property Get Headers() As Variant
    Headers = Array( _
        "Table Name", "File Name", _
        "Worksheet Name", "External Table Name", _
        "Skip")
End Property ' Headers

Public Property Get Dict() As Dictionary
   Set Dict = This.Dict
End Property ' Dict

Public Property Get SpecificTable() As ListObject
    ' Table in this workbook
    Set SpecificTable = TableBasicsSheet.ListObjects("TableBasicsTable")
End Property ' SpecificTable

Public Property Get Initialized() As Boolean
   Initialized = This.Initialized
End Property ' Initialized

Public Sub Initialize()

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    Dim LocalTable As TableBasics_Table
    Set LocalTable = New TableBasics_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(LocalTable, This.Dict, TableBasics.SpecificTable) Then
        This.Initialized = True
    Else
        ReportError "Error copying TableBasics table", "Routine", RoutineName
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
End Sub ' TableBasicsInitialize

Public Sub Reset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub ' Reset

Public Property Get HeaderWidth() As Long
    HeaderWidth = pHeaderWidth
End Property ' HeaderWidth

Public Function CreateKey(ByVal Record As TableBasics_Table) As String

    Const RoutineName As String = Module_Name & "CreateKey"
    On Error GoTo ErrorHandler

    CreateKey = Record.TableName

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
        ReportError "Error copying %1 dictionary to array,", "Routine", RoutineName
        TryCopyDictionaryToArray = False
        GoTo Done
    End If

    Dim I As Long
    I = 1

    Dim Record As TableBasics_Table
    Dim Entry As Variant
    For Each Entry In Dict.Keys
        Set Record = Dict.Item(Entry)

        Ary(I, pTableNameColumn) = Record.TableName
        Ary(I, pFileNameColumn) = Record.FileName
        Ary(I, pWorksheetNameColumn) = Record.WorksheetName
        Ary(I, pExternalTableNameColumn) = Record.ExternalTableName
        Ary(I, pSkipColumn) = Record.Skip

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
End Function ' TableBasicsTryCopyDictionaryToArray

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
    Dim Record As TableBasics_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Set Record = New TableBasics_Table

            Record.TableName = Ary(I, pTableNameColumn)
            Record.FileName = Ary(I, pFileNameColumn)
            Record.WorksheetName = Ary(I, pWorksheetNameColumn)
            Record.ExternalTableName = Ary(I, pExternalTableNameColumn)
            Record.Skip = Ary(I, pSkipColumn)

            Key = TableBasics.CreateKey(Record)

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
End Function ' TableBasicsTryCopyArrayToDictionary

Public Sub FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "TableBasicsFormatArrayAndWorksheet"
    On Error GoTo ErrorHandler


Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' TableBasicsFormatArrayAndWorksheet

' No application unique routines found

