VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HuEmployeeDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Id As Long
Public Code As String
Public FullName As String
Public BirthDate As Date
Public Address As String
Public Phone As String
Public PositionId As Long
Public PositionName As String
Public BirthDateStr As String
Public Mode As Integer


