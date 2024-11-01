VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HuEmployeeRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Implements IHuEmployeeRepository
'Dim db As DBContext
Public Function GetAllEmployee(param As String) As Collection
    Dim db As DBContext
    Set db = New DBContext
    Dim employees As Collection
    Set employees = New Collection
    
    db.OpenConnection

    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = db.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "GET_ALL_EMPLOYEE_BY_KEY"
        .Parameters.Append .CreateParameter("KeySearch", adVarChar, adParamInput, 50, param)
    End With

    Set rs = cmd.Execute

    Do While Not rs.EOF
        Dim e As HuEmployeeDTO
        Set e = New HuEmployeeDTO
        e.Code = rs!Code
        e.PositionName = rs!PositionName
        e.FullName = rs!FullName
        e.BirthDate = rs!BirthDate
        e.Phone = rs!Phone
        e.Address = rs!Address
        
        employees.Add e
        rs.MoveNext
    Loop

    rs.Close
    db.CloseConnection
    
    Set GetAllEmployee = employees
End Function

Public Function Action(request As HuEmployeeDTO) As String
    On Error GoTo ErrorHandler
    Dim db As New DBContext
    Set db = New DBContext
    db.OpenConnection

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Dim rs As ADODB.Recordset

    With cmd
        .ActiveConnection = db.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "sp_ActionToDB"
        .Parameters.Append .CreateParameter("Code", adVarChar, adParamInput, 255, request.Code)
        .Parameters.Append .CreateParameter("PositionId", adInteger, adParamInput, , request.PositionId)
        .Parameters.Append .CreateParameter("FullName", adVarChar, adParamInput, 255, request.FullName)
        .Parameters.Append .CreateParameter("BirthDate", adVarChar, adParamInput, 255, request.BirthDate)
        .Parameters.Append .CreateParameter("Phone", adVarChar, adParamInput, 255, request.Phone)
        .Parameters.Append .CreateParameter("Address", adVarChar, adParamInput, 255, request.Address)
        .Parameters.Append .CreateParameter("Mode", adInteger, adParamInput, , request.Mode)
    End With

    Set rs = cmd.Execute
    Action = "Ok"
    Exit Function
ErrorHandler:
    Action = "Error: " & Err.Description
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set cmd = Nothing
    db.CloseConnection
    Set db = Nothing
    

End Function
