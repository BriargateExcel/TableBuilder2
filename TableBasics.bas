Attribute VB_Name = "TableBasics"
Option Explicit

' Built on 3/15/2020 10:39:08 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableBasics."

Private Type TableBasicsType
    Initialized As Boolean
    Dict As Dictionary
End Type

Private This As TableBasicsType

' No application specific declarations found

Private Const pTableNameColumn As Long = 1
Private Const pFileNameColumn As Long = 2
Private Const pWorksheetNameColumn As Long = 3
Private Const pExternalTableNameColumn As Long = 4
Private Const pHeaderWidth As Long = 4

Public Property Get TableBasicsTableNameColumn() As Long
    TableBasicsTableNameColumn = pTableNameColumn
End Property

Public Property Get TableBasicsFileNameColumn() As Long
    TableBasicsFileNameColumn = pFileNameColumn
End Property

Public Property Get TableBasicsWorksheetNameColumn() As Long
    TableBasicsWorksheetNameColumn = pWorksheetNameColumn
End Property

Public Property Get TableBasicsExternalTableNameColumn() As Long
    TableBasicsExternalTableNameColumn = pExternalTableNameColumn
End Property

Public Property Get TableBasicsHeaders() As Variant
    TableBasicsHeaders = Array( _
        "Table Name", _
        "File Name", "Worksheet Name", _
        "External Table Name")
End Property

Public Property Get TableBasicsDictionary() As Dictionary
   Set TableBasicsDictionary = This.Dict
End Property

Public Property Get TableBasicsTable() As ListObject
    Set TableBasicsTable = TableBasicsSheet.ListObjects("TableBasicsTable")
End Property

Public Property Get TableBasicsInitialized() As Boolean
   TableBasicsInitialized = This.Initialized
End Property

Public Sub TableBasicsInitialize()

    Const RoutineName As String = Module_Name & "TableBasicsInitialize"
    On Error GoTo ErrorHandler
    Dim TableBasics As TableBasics_Table
    Set TableBasics = New TableBasics_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(TableBasics, TableBasicsTable, This.Dict) Then
        This.Initialized = True
    Else
        ReportError "Error copying TableBasics table", "Routine", RoutineName
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
End Sub ' TableBasicsInitialize

Public Sub TableBasicsReset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub

Public Property Get TableBasicsHeaderWidth() As Long
    TableBasicsHeaderWidth = pHeaderWidth
End Property

Public Function TableBasicsTryCopyDictionaryToArray( _
    ByVal Dict As Dictionary, _
    ByRef Ary As Variant _
    ) As Boolean

    Const RoutineName As String = Module_Name & "TableBasicsTryCopyDictionaryToArray"
    On Error GoTo ErrorHandler

    TableBasicsTryCopyDictionaryToArray = True

    If Dict.Count = 0 Then
        ReportError "Error copying TableBasics dictionary to array,", "Routine", RoutineName
        TableBasicsTryCopyDictionaryToArray = False
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

Public Function TableBasicsTryCopyArrayToDictionary( _
       ByVal Ary As Variant, _
       ByRef Dict As Dictionary _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TableBasicsTryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TableBasicsTryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record As TableBasics_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Key = Ary(I, 1)

            If Dict.Exists(Key) Then
                ReportWarning "Duplicate key", "Routine", RoutineName, "Key", Key
                TableBasicsTryCopyArrayToDictionary = False
                GoTo Done
            Else
                Set Record = New TableBasics_Table

                Record.TableName = Ary(I, pTableNameColumn)
                Record.FileName = Ary(I, pFileNameColumn)
                Record.WorksheetName = Ary(I, pWorksheetNameColumn)
                Record.ExternalTableName = Ary(I, pExternalTableNameColumn)

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
End Function ' TableBasicsTryCopyArrayToDictionary

Public Sub TableBasicsFormatArrayAndWorksheet( _
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

