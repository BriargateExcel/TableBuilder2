VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAModuleList_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iTable

' Built on 5/22/2020 4:02:54 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private Type PrivateType
    Module As String
    Extra As String
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Public Property Get Module() As String
    Module = This.Module
End Property ' Module

Public Property Let Module(ByVal Param As String)
    This.Module = Param
End Property ' Module

Public Property Get Extra() As String
    Extra = This.Extra
End Property ' Extra

Public Property Let Extra(ByVal Param As String)
    This.Extra = Param
End Property ' Extra

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = VBAModuleList.Dict
End Property ' LocalDictionary

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = VBAModuleList.HeaderWidth
End Property ' HeaderWidth

Public Property Get iTable_Headers() As Variant
    iTable_Headers = VBAModuleList.Headers
End Property ' Headers

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = VBAModuleList.Initialized
End Property ' Initialized

Public Sub iTable_Initialize()
    VBAModuleList.Initialize
End Sub ' Initialize

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = VBAModuleList.SpecificTable
End Property ' LocalTable

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "VBAModuleList_Table"
End Property ' LocalName

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = VBAModuleList.TryCopyArrayToDictionary(Ary, Dict)
End Function ' TryCopyArrayToDictionary

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = VBAModuleList.TryCopyDictionaryToArray(Dict, Ary)
End Function ' TryCopyDictionaryToArray

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    VBAModuleList.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = VBAModuleList.CreateKey(Record)
End Property ' CreateKey

Public Property Get iTable_IsDatabase() As Boolean
    iTable_IsDatabase = False
End Property ' IsDictionary

Public Property Get iTable_DatabaseName() As String
    iTable_DatabaseName = ""
End Property ' DatabaseName

Public Property Get iTable_DatabaseTableName() As String
    iTable_DatabaseTableName = "VBAModuleList"
End Property ' DatabaseTableName

' No application unique routines found

