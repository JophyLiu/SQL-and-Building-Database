VERSION 5.00
Begin VB.Form 删除页面 
   BackColor       =   &H00C0C0C0&
   Caption         =   "删除信息"
   ClientHeight    =   4500
   ClientLeft      =   5895
   ClientTop       =   2895
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8850
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox id_text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "删除的员工编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   8880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "删除员工信息"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "删除页面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim cn As ADODB.Connection
Dim strCnn As String, strSQL As String
Dim PersonID As String
Dim id As Integer
PersonID = Trim(id_text.Text) '删除的ID号
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Set cn = New ADODB.Connection
cn.Open strCnn
cn.Errors.Clear
On Error GoTo Error11
strSQL = "delect from Staff_basic_info where staff_number=" + PersonID
cn.Execute strSQL, RecordsAffected, adCmdText
If RecordsAffected <> 0 Then
MsgBox ("删除员工信息成功")
Else
MsgBox ("删除员工信息失败")
End If
cn.Close
Set cn = Nothing
Exit Sub
Error11:
 MsgBox ("删除员工信息失败")
cn.Close
Set cn = Nothing
End Sub

Private Sub Command2_Click()
Unload Me
Load 功能页面
End Sub

Private Sub Form_Load()
Me.Show
End Sub
