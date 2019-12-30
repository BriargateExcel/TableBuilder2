VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDetails_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iTable

' Built on 12/30/2019 7:59:21 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details


Private pColumnHeader As String
Private pVariableName As String
Private pFormatted As Boolean
Private pVariableType As String

Public Property Get ColumnHeader() As String
    ColumnHeader = pColumnHeader
End Property

Public Property Let ColumnHeader(ByVal Param As String)
    pColumnHeader = Param
End Property

Public Property Get VariableName() As String
    VariableName = pVariableName
End Property

Public Property Let VariableName(ByVal Param As String)
    pVariableName = Param
End Property

Public Property Get Formatted() As Boolean
    Formatted = pFormatted
End Property

Public Property Let Formatted(ByVal Param As Boolean)
    pFormatted = Param
End Property

Public Property Get VariableType() As String
    VariableType = pVariableType
End Property

Public Property Let VariableType(ByVal Param As String)
    pVariableType = Param
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableDetails.Headers
End Property

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableDetailsDictionary
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableDetailsHeaderWidth
End Property


Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableDetailsInitialized
End Property


Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableDetailsTable
End Property


Public Sub iTable_Initialize()
    TableDetails.Initialize
End Sub


Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableDetails.TryCopyArrayToDictionary(Ary, Dict)
End Function


Public Sub iTable_CopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant)
    TableDetails.CopyDictionaryToArray Dict, Ary
End Sub


Public Function TryCopyDictionaryToTable( _
       Optional ByVal Dict As Dictionary = Nothing, _
       Optional ByVal Table As ListObject = Nothing, _
       Optional TableCorner As Range = Nothing, _
       Optional TableName As String)


    TryCopyDictionaryToTable = TableDetails.TryCopyDictionaryToTable(Dict, Table, TableCorner, TableName)
End Function

