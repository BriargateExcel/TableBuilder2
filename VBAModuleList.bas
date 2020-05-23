Attribute VB_Name = "VBAModuleList"
Option Explicit

' Built on 5/22/2020 4:05:43 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "VBAModuleList."

Private Type PrivateType
    Initialized As Boolean
    Dict As Dictionary
    Wkbk As Workbook
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Private Const pModuleColumn As Long = 1
Private Const pExtraColumn As Long = 2
Private Const pHeaderWidth As Long = 2

Private Const pFileName As String = "Blank"
Private Const pWorksheetName As String = vbNullString
Private Const pExternalTableName As String = vbNullString

Public Property Get ModuleColumn() As Long
    ModuleColumn = pModuleColumn
End Property ' ModuleColumn

Public Property Get ExtraColumn() As Long
    ExtraColumn = pExtraColumn
End Property ' ExtraColumn

Public Property Get Headers() As Variant
    Headers = Array( _
        "Module", _
        "Extra")
End Property ' Headers

Public Property Get Dict() As Dictionary
   Set Dict = This.Dict
End Property ' Dict

Public Property Get SpecificTable() As ListObject
    ' This table is handled in other ways
    Set SpecificTable = Nothing
End Property ' SpecificTable

Public Property Get Initialized() As Boolean
   Initialized = This.Initialized
End Property ' Initialized

Public Sub Initialize()

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    Dim LocalTable As VBAModuleList_Table
    Set LocalTable = New VBAModuleList_Table

    Set This.Dict = New Dictionary
    If Table.TryCopyTableToDictionary(LocalTable, This.Dict, VBAModuleList.SpecificTable) Then
        This.Initialized = True
    Else
        ReportError "Error copying VBAModuleList table", "Routine", RoutineName
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
End Sub ' VBAModuleListInitialize

Public Sub Reset()
    This.Initialized = False
    Set This.Dict = Nothing
End Sub ' Reset

Public Property Get HeaderWidth() As Long
    HeaderWidth = pHeaderWidth
End Property ' HeaderWidth

Public Function CreateKey(ByVal Record As VBAModuleList_Table) As String

    Const RoutineName As String = Module_Name & "CreateKey"
    On Error GoTo ErrorHandler

    CreateKey = Record.Module

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
        ReportError "Error copying VBAModuleList_Table dictionary to array,", "Routine", RoutineName
        TryCopyDictionaryToArray = False
        GoTo Done
    End If

    ReDim Ary(1 To Dict.Count, 1 To 2)

    Dim I As Long
    I = 1

    Dim Record As VBAModuleList_Table
    Dim Entry As Variant
    For Each Entry In Dict.Keys
        Set Record = Dict.Item(Entry)

        Ary(I, pModuleColumn) = Record.Module
        Ary(I, pExtraColumn) = Record.Extra

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
End Function ' VBAModuleListTryCopyDictionaryToArray

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
    Dim Record As VBAModuleList_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Set Record = New VBAModuleList_Table

            Record.Module = Ary(I, pModuleColumn)
            Record.Extra = Ary(I, pExtraColumn)

            Key = VBAModuleList.CreateKey(Record)

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
End Function ' VBAModuleListTryCopyArrayToDictionary

Public Sub FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "VBAModuleListFormatArrayAndWorksheet"
    On Error GoTo ErrorHandler


Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description

    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' VBAModuleListFormatArrayAndWorksheet

' No application unique routines found

