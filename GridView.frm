VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLearn2 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExport 
      BackColor       =   &H00FFFF00&
      Caption         =   "Export"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   7920
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnEdit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Edit"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton btnAdd 
      BackColor       =   &H0080FF80&
      Caption         =   "Add"
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid grdEmployee 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12303
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLearn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As DBContext
Dim employeeRepo As HuEmployeeRepository
Private SelectData As HuEmployeeDTO
Public Sub HoldingData(ByRef data As HuEmployeeDTO)
    Set SelectData = data
End Sub
Private Sub btnAdd_Click()
    frmPageEdit.Show
End Sub

Private Sub btnDelete_Click()
    If Not SelectData Is Nothing Then
        Dim result As Integer

        result = MsgBox("You want to delete record?", vbYesNo + vbQuestion, "Ok")
        
        If result = vbYes Then
        Dim employeeRepo As New HuEmployeeRepository
        'Set employeeRepo = New HuEmployeeRepository
            Dim str As String
            SelectData.Mode = 0
            str = employeeRepo.Action(SelectData)
            Dim allEmployees As Collection
            Set allEmployees = employeeRepo.GetAllEmployee(txtSearch.Text)

            Set grdEmployee.DataSource = ConvertColectionToDatatable(allEmployees)
            MsgBox str
        End If
    Else
        MsgBox "No position selected."
    End If
End Sub

Private Sub btnEdit_Click()
    frmPageEdit.Show
    frmPageEdit.SetPositionData SelectData
End Sub
Private Sub ExportEmployeesToCSV(filePath As String, employees As Collection)
     Dim fileNum As Integer
    Dim i As Integer
    Dim csvLine As String
    Dim employee As HuEmployeeDTO

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    Print #fileNum, "Id,PositionName, Code,FullName, BirthDate, Phone, Address"

    For i = 1 To employees.Count
        Set employee = employees(i)
        
        csvLine = employee.Id & "," & employee.PositionName & "," & _
        employee.Code & "," & employee.FullName & "," & employee.BirthDate & "," & employee.Phone & "," & employee.Address
        
        Print #fileNum, csvLine
    Next i

    Close #fileNum

    MsgBox "Xu?t d? li?u nhân viên thành công vào file CSV!"
End Sub
Private Sub btnExport_Click()
    Dim employeeRepo As HuEmployeeRepository
    Set employeeRepo = New HuEmployeeRepository

    Dim allEmployees As Collection
    Set allEmployees = employeeRepo.GetAllEmployee(txtSearch.Text)

    Dim filePath As String
    filePath = "C:\Users\datnv02\Export\exportdata.csv" ' Tam thoi fix cung path
    ExportEmployeesToCSV filePath, allEmployees
End Sub

Private Sub Form_Load()
    Dim employeeRepo As HuEmployeeRepository
    Set employeeRepo = New HuEmployeeRepository
    
    Dim allEmployees As Collection
    Set allEmployees = employeeRepo.GetAllEmployee("")

    Set grdEmployee.DataSource = ConvertColectionToDatatable(allEmployees)

End Sub


Private Sub txtSearch_Change()
    Dim employeeRepo As HuEmployeeRepository
    Set employeeRepo = New HuEmployeeRepository

    Dim allEmployees As Collection
    Set allEmployees = employeeRepo.GetAllEmployee(txtSearch.Text)

    Set grdEmployee.DataSource = ConvertColectionToDatatable(allEmployees)
    grdEmployee.Columns(6).Visible = False
End Sub
Public Function ConvertColectionToDatatable(listData As Collection) As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    rs.Fields.Append "Code", adVarChar, 50
    rs.Fields.Append "PositionName", adVarChar, 50
    rs.Fields.Append "FullName", adVarChar, 100
    rs.Fields.Append "BirthDate", adDate
    rs.Fields.Append "Phone", adVarChar, 20
    rs.Fields.Append "Address", adVarChar, 255
    rs.Fields.Append "PositionId", adInteger
    
    rs.Open

    Dim employee As HuEmployeeDTO
    Dim i As Integer
    For i = 1 To listData.Count
        Set employee = listData(i)
        rs.AddNew
        rs.Fields("Code").Value = employee.Code
        rs.Fields("PositionName").Value = employee.PositionName
        rs.Fields("FullName").Value = employee.FullName
        rs.Fields("BirthDate").Value = employee.BirthDate
        rs.Fields("Phone").Value = employee.Phone
        rs.Fields("Address").Value = employee.Address
        rs.Fields("PositionId").Value = employee.PositionId
        rs.Update
    Next i

    Set ConvertColectionToDatatable = rs
End Function
Private Sub grdEmployee_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim selectedRow As Integer
    Dim obj As HuEmployeeDTO
 
    Set obj = New HuEmployeeDTO
    selectedRow = grdEmployee.Row
    
    obj.Code = CInt(grdEmployee.Columns(0).Text)
    obj.PositionName = grdEmployee.Columns(1).Value
    obj.FullName = grdEmployee.Columns(2).Text
    obj.BirthDate = grdEmployee.Columns(3).Text
    obj.Phone = grdEmployee.Columns(4).Text
    obj.Address = grdEmployee.Columns(5).Text
    obj.PositionId = grdEmployee.Columns(6).Value

    HoldingData obj
End Sub
