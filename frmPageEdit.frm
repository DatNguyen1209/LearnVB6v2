VERSION 5.00
Begin VB.Form frmPageEdit 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   840
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtBirthDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   6255
   End
   Begin VB.TextBox txtPhone 
      Height          =   405
      Left            =   5760
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtFullName 
      Height          =   405
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox cbxListPosition 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Address:"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Phone:"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Birthdate: "
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Full name:"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Position:"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DBContext


Private Sub btnCancel_Click()
    frmPageEdit.Hide
End Sub

Private Sub btnSave_Click()
    Dim employeeRepo As New HuEmployeeRepository
    'Set employeeRepo = New HuEmployeeRepository
    Dim result As String
    Dim request As HuEmployeeDTO
    Set request = New HuEmployeeDTO
    Dim PositionId As Long
    PositionId = cbxListPosition.ItemData(cbxListPosition.ListIndex)
    request.FullName = txtFullName.Text
    request.BirthDate = txtBirthDate.Text
    request.Phone = txtPhone.Text
    request.PositionId = PositionId
    request.Address = txtAddress.Text
    'Create
    If txtEdit = "1" Then
        request.Code = txtCode.Text
        request.Mode = 2
    Else
        request.Mode = 1
    End If
    result = employeeRepo.Action(request)
    frmPageEdit.Hide
    Dim allEmployees As Collection
    Set allEmployees = employeeRepo.GetAllEmployee(frmLearn2.txtSearch.Text)

    Set frmLearn2.grdEmployee.DataSource = frmLearn2.ConvertColectionToDatatable(allEmployees)
End Sub

Private Sub Form_Load()
     Dim db As New DBContext
    Dim rs As ADODB.Recordset
    Dim query As String
    db.OpenConnection

    query = "SELECT ID AS Id, NAME AS Name FROM POSITION"

    Set rs = db.ExecuteQuery(query)

    cbxListPosition.Clear

    With cbxListPosition
    Do While Not rs.EOF
        .AddItem rs.Fields(1)
        .ItemData(.NewIndex) = rs.Fields(0)
        rs.MoveNext
        Loop
    End With

    rs.Close

    db.CloseConnection
    txtBirthDate.Text = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
        Dim currentText As String
        currentText = txtBirthDate.Text
        Select Case Len(currentText)
            Case 2
                txtBirthDate.Text = currentText & "/"
                txtBirthDate.SelStart = Len(txtBirthDate.Text)
            Case 5
                txtBirthDate.Text = currentText & "/"
                txtBirthDate.SelStart = Len(txtBirthDate.Text)
        End Select
    ElseIf KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Public Sub SetPositionData(ByRef dataIn As HuEmployeeDTO)
    txtEdit.Text = "1"
    txtCode.Text = dataIn.Code
    txtFullName.Text = dataIn.FullName
    txtAddress.Text = dataIn.Address
    txtBirthDate.Text = dataIn.BirthDate
    txtPhone.Text = dataIn.Phone
    txtAddress.Text = dataIn.Address
    'cbxListPosition.DataChanged = dataIn.PositionName
    For i = 0 To cbxListPosition.ListCount - 1
        If cbxListPosition.List(i) = dataIn.PositionName Then
            cbxListPosition.ListIndex = i ' L?y key tuong ?ng
            
            Exit For
        End If
    Next i
End Sub
