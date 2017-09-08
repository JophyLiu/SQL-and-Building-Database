VERSION 5.00
Begin VB.Form 登记页面 
   BackColor       =   &H00C0C0C0&
   Caption         =   "登记信息"
   ClientHeight    =   10290
   ClientLeft      =   4455
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   13005
   Begin VB.TextBox Text31 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   68
      Top             =   9480
      Width           =   2295
   End
   Begin VB.TextBox Text30 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   66
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox Text29 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2880
      TabIndex        =   65
      Top             =   9240
      Width           =   2055
   End
   Begin VB.TextBox Text28 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   62
      Top             =   8880
      Width           =   2295
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   7320
      TabIndex        =   61
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   60
      Top             =   8760
      Width           =   2055
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2880
      TabIndex        =   59
      Top             =   8280
      Width           =   2055
   End
   Begin VB.TextBox Text24 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   53
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   52
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   2880
      TabIndex        =   49
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   2880
      TabIndex        =   48
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   2880
      TabIndex        =   45
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   43
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   41
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   7320
      TabIndex        =   39
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   7320
      TabIndex        =   38
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox Text14 
      Height          =   390
      Left            =   2880
      TabIndex        =   35
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox Text13 
      Height          =   390
      Left            =   2880
      TabIndex        =   34
      Top             =   4560
      Width           =   2175
   End
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
      Left            =   11280
      TabIndex        =   29
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "入库"
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
      Left            =   11280
      TabIndex        =   28
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   10080
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8640
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   3120
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "女"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "男"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0C0C0&
      Caption         =   "备注"
      Height          =   375
      Left            =   6720
      TabIndex        =   67
      Top             =   9480
      Width           =   975
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0C0C0&
      Caption         =   "调入时间"
      Height          =   255
      Left            =   2040
      TabIndex        =   64
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0C0C0&
      Caption         =   "调出时间"
      Height          =   255
      Left            =   2040
      TabIndex        =   63
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0C0C0&
      Caption         =   "新职务"
      Height          =   375
      Left            =   6600
      TabIndex        =   58
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0C0C0&
      Caption         =   "原职务"
      Height          =   375
      Left            =   6600
      TabIndex        =   57
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "新部门"
      Height          =   255
      Left            =   2160
      TabIndex        =   56
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0C0C0&
      Caption         =   "原部门"
      Height          =   735
      Left            =   2160
      TabIndex        =   55
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C0C0&
      Caption         =   "录入员工调动信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   4680
      TabIndex        =   54
      Top             =   7680
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   12960
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0C0C0&
      Caption         =   "出差开始时间"
      Height          =   375
      Left            =   6120
      TabIndex        =   51
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0C0C0&
      Caption         =   "出差天数"
      Height          =   375
      Left            =   6480
      TabIndex        =   50
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0C0&
      Caption         =   " 加班日期"
      Height          =   375
      Left            =   1920
      TabIndex        =   47
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0C0C0&
      Caption         =   " 加班天数"
      Height          =   375
      Left            =   1920
      TabIndex        =   46
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      Caption         =   "假期开始时间"
      Height          =   375
      Left            =   1680
      TabIndex        =   44
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "事假天数"
      Height          =   495
      Left            =   6480
      TabIndex        =   42
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "病假天数"
      Height          =   375
      Left            =   6480
      TabIndex        =   40
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "早退次数"
      Height          =   255
      Left            =   6480
      TabIndex        =   37
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "迟到次数"
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "下班时间"
      Height          =   375
      Left            =   2040
      TabIndex        =   33
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "上班时间"
      Height          =   375
      Left            =   2040
      TabIndex        =   32
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "录入员工考勤信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   4680
      TabIndex        =   31
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   12960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "录入员工基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   4560
      TabIndex        =   30
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "      专业"
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "      学历"
      Height          =   255
      Left            =   6240
      TabIndex        =   26
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  E-mail"
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   住址"
      Height          =   375
      Left            =   2040
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "     年龄"
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "日"
      Height          =   255
      Left            =   11160
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "月"
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   " 年"
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  出生日期"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "进入本单位时间"
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  性别"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  编号"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  籍贯"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "  姓名"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "登记页面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public String_Area As String              '地区编号
Public Name_Area As String                '地区名称
Public Num_Archive As String              '成果个数
Public Sum_People_Area                    '某地区人数
Public ID_Person As String
Public ByName As String
Public strCnn As String
Public conODBC As ADODB.Connection
Public Sub Sub_ConnectServer()
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Set conODBC = New ADODB.Connection
conODBC.Open strCnn
End Sub
Private Sub Command1_Click()
'-----------------基本信息表-------------
staff_n = Trim(Text3.Text)
If Option1.Value = True Then
sex = "男"
Else
sex = "女"
End If
num = Trim(Text3.Text)
jiguan = Trim(Text2.Text)
addr = Trim(Text9.Text)
birth = Trim(Text5.Text) + "-" + Trim(Text6.Text) + "-" + Trim(Text7.Text)
age = Trim(Text8.Text)
em = Trim(Text10.Text)
xueli = Trim(Text11.Text)
zhuanye = Trim(Text12.Text)
t = Trim(Text4.Text)
'----------------考勤表------------------

st = Text13.Text  '上班时间
xt = Text14.Text '下班时间
cc = Text15.Text '迟到次数
zc = Trim(Text16.Text) '早退次数
'jc=Trim(Text17.Text) 进出标志，改改数据库的数据类型！！！！
bj = Trim(Text18.Text) '病假天数
sj = Trim(Text19.Text) '事假天数
jk = Trim(Text20.Text) '假期开始时间
jt = Trim(Text21.Text) '加班天数
jc = Trim(Text22.Text) '加班日期
ct = Trim(Text23.Text) '出差天数
ck = Trim(Text24.Text) '出差开始时间
'----------------------员工调动表---------------
ybm = Trim(Text25.Text)
jbm = Trim(Text26.Text)
yzw = Trim(Text27.Text)
xzw = Trim(Text28.Text)
dc = Trim(Text29.Text)
dr = Trim(Text30.Text)
info = Trim(Text31.Text)
'====================录入信息数据============================================
'====================用ADO方式代开数据库=====================================
Sub_ConnectServer

conODBC.BeginTrans
Set rs = New ADODB.Recordset
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
'-----------------录入基本信息表---------------------------------
rs.Open "Select * From staff_basic_info ", conODBC, adCmdTable
rs.MoveLast
rs.AddNew
rs!staff_name = staff_n
rs!staff_number = num
rs!staff_sex = sex
rs!staff_where = jiguan
rs!staff_age = age
rs!staff_birth = birth
rs!staff_add = addr
rs!staff_Email = em
rs!staff_ROFS = xueli
rs!staff_major = zhuanye
rs!staff_intime = t
rs.Update
rs.Close
'------------录入员工考勤信息---------------------------------------
rs.Open "Select * From staff_attend_info ", conODBC, adCmdTable
rs.MoveLast
rs.AddNew
rs!staff_number = num
rs!go_work_time = st
rs!out_work_time = xt
rs!late_times = cc
rs!leave_early_times = zc
rs!sicks = bj
rs!affair = sj
rs!leaves_start = jk
rs!work_overtime = jt
rs!overtime_data = jc
rs!business_trip = ct
rs!B_trip_start = ck
rs.Update
rs.Close
'---------------录入调动信息-----------------------------------
rs.Open "Select * From staff_attend_info ", conODBC, adCmdTable
rs.MoveLast
rs.AddNew
rs!old_department = ybm
rs!new_departmeent = jbm
rs!old_position = yzw
rs!new_position = xzw
rs!out_data = dc
rs!in_data = dr
rs!info = info
rs.Update
rs.Close
conODBC.CommitTrans
MsgBox "入库成功！", vbOKOnly, "信息提示"
End Sub

Private Sub Command2_Click()
Unload Me
Load 功能页面
End Sub

Private Sub Form_Load()
Me.Show
End Sub

