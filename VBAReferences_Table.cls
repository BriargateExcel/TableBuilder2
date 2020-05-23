VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAReferences_Table"
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
    Name As String
    Description As String
    GUID As String
    Major As String
    Minor As String
End Type ' PrivateType

Private This As PrivateType

' No application specific declarations found

Public Property Get Name() As String
    Name = This.Name
End Property ' Name

Public Property Let Name(ByVal Param As String)
    This.Name = Param
End Property ' Name

Public Property Get Description() As String
    Description = This.Description
End Property ' Description

Public Property Let Description(ByVal Param As String)
    This.Description = Param
End Property ' Description

Public Property Get GUID() As String
    GUID = This.GUID
End Property ' GUID

Public Property Let GUID(ByVal Param As String)
    This.GUID = Param
End Property ' GUID

Public Property Get Major() As String
    Major = This.Major
End Property ' Major

Public Property Let Major(ByVal Param As String)
    This.Major = Param
End Property ' Major

Public Property Get Minor() As String
    Minor = This.Minor
End Property ' Minor

Public Property Let Minor(ByVal Param As String)
    This.Minor = Param
End Property ' Minor

Public Property Get iTable_LocalDictionary() As Dictionary
    Set iTable_LocalDictionary = VBAReferences.Dict
End Property ' LocalDictionary

Public Property Get iTable_HeaderWidth() As Long
    iTable_HeaderWidth = VBAReferences.HeaderWidth
End Property ' HeaderWidth

Public Property Get iTable_Headers() As Variant
    iTable_Headers = VBAReferences.Headers
End Property ' Headers

Public Property Get iTable_Initialized() As Boolean
    iTable_Initialized = VBAReferences.Initialized
End Property ' Initialized

Public Sub iTable_Initialize()
    VBAReferences.Initialize
End Sub ' Initialize

Public Property Get iTable_LocalTable() As ListObject
    Set iTable_LocalTable = VBAReferences.SpecificTable
End Property ' LocalTable

Public Property Get iTable_LocalName() As String
    iTable_LocalName = "VBAReferences_Table"
End Property ' LocalName

Public Function iTable_TryCopyArrayToDictionary(ByVal Ary As Variant, ByRef Dict As Dictionary) As Boolean
    iTable_TryCopyArrayToDictionary = VBAReferences.TryCopyArrayToDictionary(Ary, Dict)
End Function ' TryCopyArrayToDictionary

Public Function iTable_TryCopyDictionaryToArray(ByVal Dict As Dictionary, ByRef Ary As Variant) As Boolean
    iTable_TryCopyDictionaryToArray = VBAReferences.TryCopyDictionaryToArray(Dict, Ary)
End Function ' TryCopyDictionaryToArray

Public Sub iTable_FormatArrayAndWorksheet( _
    ByRef Ary As Variant, _
    ByVal Table As ListObject)
    VBAReferences.FormatArrayAndWorksheet Ary, Table
End Sub ' FormatArrayAndWorksheet

Public Property Get iTable_CreateKey(ByVal Record As iTable) As String
    iTable_CreateKey = VBAReferences.CreateKey(Record)
End Property ' CreateKey

Public Property Get iTable_IsDatabase() As Boolean
    iTable_IsDatabase = False
End Property ' IsDictionary

Public Property Get iTable_DatabaseName() As String
    iTable_DatabaseName = ""
End Property ' DatabaseName

Public Property Get iTable_DatabaseTableName() As String
    iTable_DatabaseTableName = "VBAReferences"
End Property ' DatabaseTableName

' No application unique routines found

