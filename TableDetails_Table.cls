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

' Built on 6/12/2020 3:16:22 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Type PrivateType
    ColumnHeader As String
    VariableName As String
    VariableType As String
    Key As String
    Format As String
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Public Property Get ColumnHeader() As String
    ColumnHeader = This.ColumnHeader
End Property ' ColumnHeader

Public Property Let ColumnHeader(ByVal Param As String)
    This.ColumnHeader = Param
End Property ' ColumnHeader

Public Property Get VariableName() As String
    VariableName = This.VariableName
End Property ' VariableName

Public Property Let VariableName(ByVal Param As String)
    This.VariableName = Param
End Property ' VariableName

Public Property Get VariableType() As String
    VariableType = This.VariableType
End Property ' VariableType

Public Property Let VariableType(ByVal Param As String)
    This.VariableType = Param
End Property ' VariableType

Public Property Get Key() As String
    Key = This.Key
End Property ' Key

Public Property Let Key(ByVal Param As String)
    This.Key = Param
End Property ' Key

Public Property Get Format() As String
    Format = This.Format
End Property ' Format

Public Property Let Format(ByVal Param As String)
    This.Format = Param
End Property ' Format

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = TableDetails.Dict
End Property ' LocalDictionary

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableDetails.HeaderWidth
End Property ' HeaderWidth

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableDetails.Headers
End Property ' Headers

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableDetails.Initialized
End Property ' Initialized

Public Sub iTable_Initialize()
    TableDetails.Initialize
End Sub ' Initialize

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableDetails.SpecificTable
End Property ' LocalTable

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableDetails_Table"
End Property ' LocalName

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableDetails.TryCopyArrayToDictionary(Ary, Dict)
End Function ' TryCopyArrayToDictionary

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableDetails.TryCopyDictionaryToArray(Dict, Ary)
End Function ' TryCopyDictionaryToArray

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableDetails.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = TableDetails.CreateKey(Record)
End Property ' CreateKey

Public Property Get iTable_IsDatabase() As Boolean
    iTable_IsDatabase = False
End Property ' IsDictionary

Public Property Get iTable_DatabaseName() As String
    iTable_DatabaseName = ""
End Property ' DatabaseName

Public Property Get iTable_DatabaseTableName() As String
    iTable_DatabaseTableName = "TableDetails"
End Property ' DatabaseTableName

' No application unique routines found

