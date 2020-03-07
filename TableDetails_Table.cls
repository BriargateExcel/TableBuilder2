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

' Built on 3/6/2020 11:23:29 AM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Type TableDetailsType
    ColumnHeader As String
    VariableName As String
    VariableType As String
    Key As String
    Format As String
End Type

Private This As TableDetailsType

' No application specific declarations found

Public Property Get ColumnHeader() As String
    ColumnHeader = This.ColumnHeader
End Property

Public Property Let ColumnHeader(ByVal Param As String)
    This.ColumnHeader = Param
End Property

Public Property Get VariableName() As String
    VariableName = This.VariableName
End Property

Public Property Let VariableName(ByVal Param As String)
    This.VariableName = Param
End Property

Public Property Get VariableType() As String
    VariableType = This.VariableType
End Property

Public Property Let VariableType(ByVal Param As String)
    This.VariableType = Param
End Property

Public Property Get Key() As String
    Key = This.Key
End Property

Public Property Let Key(ByVal Param As String)
    This.Key = Param
End Property

Public Property Get Format() As String
    Format = This.Format
End Property

Public Property Let Format(ByVal Param As String)
    This.Format = Param
End Property

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableDetailsDictionary
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableDetailsHeaderWidth
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableDetailsHeaders
End Property

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableDetailsInitialized
End Property

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableDetailsTable
End Property

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableDetails_Table"
End Property

Public Sub iTable_Initialize()
    TableDetailsInitialize
End Sub

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableDetailsTryCopyArrayToDictionary(Ary, Dict)
End Function

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableDetailsTryCopyDictionaryToArray(Dict, Ary)
End Function

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableDetailsFormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

' No application unique routines found

