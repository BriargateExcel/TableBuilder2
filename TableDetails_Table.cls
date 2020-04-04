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

' Built on 3/23/2020 11:18:26 AM
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
    Set iTable_LocalDictionary = TableDetails.Dict
End Property

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = TableDetails.HeaderWidth
End Property

Public Property Get iTable_Headers() As Variant
    iTable_Headers = TableDetails.Headers
End Property

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = TableDetails.Initialized
End Property

Public Sub iTable_Initialize()
    TableDetails.Initialize
End Sub

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = TableDetails.SpecificTable
End Property

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "TableDetails_Table"
End Property

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = TableDetails.TryCopyArrayToDictionary(Ary, Dict)
End Function

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = TableDetails.TryCopyDictionaryToArray(Dict, Ary)
End Function

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    TableDetails.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = TableDetails.CreateKey(Record)
End Property

' No application unique routines found

