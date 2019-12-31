VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableBasics_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Built on 12/31/2019 12:09:32 PM
' Built By Briargate Excel Table Builder
' See BriargateExcel.com for details

Private pTableName As String

Public Property Get TableName() As String
    TableName = pTableName
End Property

Public Property Let TableName(ByVal Param As String)
    pTableName = Param
End Property

