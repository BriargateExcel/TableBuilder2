Attribute VB_Name = "TableBasics"
Option Explicit

' Built on 12/30/2019 12:26:58 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Const Module_Name As String = "TableBasics."

Private pInitialized As Boolean
Private pTableBasicsDict As Dictionary

Private Const pTableNameColumn As Long = 1
Private Const pHeaderWidth As Long = 1

Public Property Get TableBasicsTable() As ListObject
    Set TableBasicsTable = TableBasicsSheet.ListObjects("TableBasicsTable")
End Property

Public Sub ResetTableBasics()
    pInitialized = False
    Set pTableBasicsDict = Nothing
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
        MsgBox "The TableBasics table is empty"
        TryCopyTableToDictionary = False
        GoTo Done
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    Dim ThisDict As Dictionary
    If Dict Is Nothing Then
        Set ThisDict = New Dictionary
    Else
        Set ThisDict = pTableBasicsDict
    End If

    If TableBasics.TryCopyArrayToDictionary(Ary, ThisDict) Then
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
    Headers = Array("TableName")
End Property

Public Property Get TableBasicsDictionary() As Dictionary
   Set TableBasicsDictionary = pTableBasicsDict
End Property

Public Property Get TableBasicsHeaderWidth() As Long
    TableBasicsHeaderWidth = pHeaderWidth
End Property

Public Property Get TableBasicsInitialized() As Boolean
    TableBasicsInitialized = pInitialized
End Property

Public Sub Initialize()

    ' This routine loads the dictionary

    Const RoutineName As String = Module_Name & "Initialize"
    On Error GoTo ErrorHandler

    pInitialized = True

    Dim TableBasics As TableBasics_Table
    Set TableBasics = New TableBasics_Table

    Set pTableBasicsDict = New Dictionary
    If Table.TryCopyTableToDictionary(TableBasics, TableBasicsTable, pTableBasicsDict) Then
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

    Dim Record As TableBasics_Table
    Dim Entry As Variant
    For Each Entry In DetailsDict.Keys
        Set Record = DetailsDict.Item(Entry)

        Ary(I, pTableNameColumn) = Record.TableName

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
    Dim Record As TableBasics_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Key = Ary(I, pTableNameColumn)

            If Dict.Exists(Key) Then
                MsgBox "Duplicate key"
                TryCopyArrayToDictionary = False
                GoTo Done
            Else
                Set Record = New TableBasics_Table

                Record.TableName = Ary(I, pTableNameColumn)

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

    Dim ClassName As TableBasics_Table
    Set ClassName = New TableBasics_Table

    '    FormatColumnAsText pFirstColumn, Table, TableCorner

    If Table.TryCopyDictionaryToTable(ClassName, Dict, Table, TableCorner, TableName) Then
        ' Success; do
    Else
        MsgBox "Error copying TableBasics dictionary to table"
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

