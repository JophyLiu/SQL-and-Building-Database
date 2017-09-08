VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form 调动信息 
   BackColor       =   &H00C0C0C0&
   Caption         =   "调动信息"
   ClientHeight    =   5475
   ClientLeft      =   5460
   ClientTop       =   2460
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9990
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   12632319
      BackColorSel    =   16777088
      ForeColorSel    =   0
      BackColorBkg    =   12632256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   9960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "员工调动信息"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "调动信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCnn As String             '数据库连接串
Public conODBC As Connection
Public Old_Row As Integer                 '浏览信息表表格前一次行数
'数据库链接对象
'-------连接数据库服务器
Public Sub Sub_ConnectServer()
Dim strCnn As String
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Dim conODBC As ADODB.Connection
Set conODBC = CreateObject("adodb.connection")
conODBC.Open strCnn
End Sub
Private Sub MSFlexGrid1_Click()
Dim no_row As Integer
no_row = MSFlexGrid1.RowSel
MSFlexGrid1.Row = Old_Row
For i = 1 To 9
MSFlexGrid1.Col = i
MSFlexGrid1.CellBackColor = vbWhite
Next
Old_Row = no_row
MSFlexGrid1.Row = no_row
For i = 1 To 9
MSFlexGrid1.Col = i
MSFlexGrid1.CellBackColor = vbYellow
Next
End Sub
Private Sub Form_Load()

MSFlexGrid1.RowHeight(0) = 500
MSFlexGrid1.FormatString = "^序号 |^员工编号 |^  原部门  |^ 新部门  |^   原职务   |^   新职务   |^  调出时间  |^  调入时间  |^  备注  "
'============用ADO方式打开数据库
'MsgBox"Opening rcgl_sys...人才管理数据库"
Dim strCnn As String
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Dim conODBC As ADODB.Connection
Set conODBC = CreateObject("adodb.connection")
conODBC.Open strCnn
 
Dim rs As ADODB.Recordset
Set rs = CreateObject("adodb.recordset")
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    rs.CursorLocation = adUseClient

    '-----------------------读人才专业信息表
    rs.Open "Select * From Staff_mobilize_info", conODBC, adOpenStatic, adLockReadOnly
    
   num_records = rs.RecordCount
   MSFlexGrid1.Rows = num_records + 1
    rs.MoveFirst
    For i = 1 To num_records
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = i
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = rs!staff_number
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Text = rs!old_department
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Text = rs!new_department
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Text = rs!old_position
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Text = rs!new_position
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = rs!out_date
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Text = rs!in_date
    MSFlexGrid1.Col = 8
    If IsNull(rs!info) Then
    MSFlexGrid1.Text = ""
    Else
    MSFlexGrid1.Text = Trim(rs!info)
    End If
  
    rs.MoveNext
   Next
    rs.Close
    Me.Show
    MSFlexGrid1.HighLight = flexHighlightWithFocus
    MSFlexGrid1.FocusRect = flexFocusheave
    MSFlexGrid1.Row = 1
    Old_Row = 1
    For i = 1 To 8 '''''''''''''''根据实际列数修改
    MSFlexGrid1.Col = i
    MSFlexGrid1.CellBackColor = vbYellow
    Next
    Me.Show
End Sub

Private Sub Command1_Click()
Unload Me
Load 功能页面
End Sub


