Attribute VB_Name = "TableBasics"
Option Explicit

' Built on 1/1/2020 9:32:28 AM
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

Public Property Get TableBasicsDictionary() As Dictionary
   Set TableBasicsDictionary = pTableBasicsDict
End Property

Public Property Get TableBasicsInitialized() As Boolean
   TableBasicsInitialized = pInitialized
End Property

Public Sub TableBasicsReset()
    pInitialized = False
    Set pTableBasicsDict = Nothing
End Sub

Public Property Get TableBasicsHeaderWidth() As Long
    TableBasicsHeaderWidth = pHeaderWidth
End Property

Public Property Get TableBasicsHeaders() As Variant
    TableBasicsHeaders = Array("Table Name")
End Property

Public Sub TableBasicsInitialize()

    Const RoutineName As String = Module_Name & "TableBasicsInitialize"
    On Error GoTo ErrorHandler

    pInitialized = True

    Dim TableBasics As TableBasics_Table
    Set TableBasics = New TableBasics_Table

    Set pTableBasicsDict = New Dictionary
    If Table.TryCopyTableToDictionary(TableBasics, TableBasicsTable, pTableBasicsDict) Then
        ' Success; do nothing
    Else
        MsgBox "Error copying TableBasics table"
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
End Sub ' TableBasicsInitialize

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
       ByRef Dict As Dictionary)

    Const RoutineName As String = Module_Name & "TableBasicsTryCopyArrayToDictionary"
    On Error GoTo ErrorHandler

    TableBasicsTryCopyArrayToDictionary = True

    Dim I As Long

    Set Dict = New Dictionary

    Dim Key As String
    Dim Record As TableBasics_Table

    If VarType(Ary) = vbArray Or VarType(Ary) = 8204 Then
        For I = 1 To UBound(Ary, 1)
            Key = Ary(I, pTableNameColumn)

            If Dict.Exists(Key) Then
                MsgBox "Duplicate key"
                TableBasicsTryCopyArrayToDictionary = False
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
End Function ' TableBasicsTryCopyArrayToDictionary

Public Sub TableBasicsFormatWorksheet(ByVal Table As ListObject)

    Const RoutineName As String = Module_Name & "TableBasicsFormatWorksheet"
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
End Sub ' TableBasicsFormatWorksheet

''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  '
'             End of Generated code                '
'            Start unique code here                '
'                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''

