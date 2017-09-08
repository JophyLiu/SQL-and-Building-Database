VERSION 5.00
Begin VB.Form 功能页面 
   BackColor       =   &H00C0C0C0&
   Caption         =   "功能选择"
   ClientHeight    =   5760
   ClientLeft      =   6180
   ClientTop       =   2175
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9255
   Begin VB.CommandButton Command6 
      Caption         =   "帮助"
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
      Left            =   6720
      TabIndex        =   9
      Top             =   4920
      Width           =   975
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
      Left            =   7920
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "管理功能"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5280
      TabIndex        =   4
      Top             =   1920
      Width           =   3015
      Begin VB.CommandButton Command9 
         Caption         =   "修改信息"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "删除信息"
         Height          =   615
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "登记信息"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "查询功能"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
      Begin VB.CommandButton Command5 
         Caption         =   "调动信息"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "条件查询"
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "考勤信息"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "基本信息"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "版本   V 1.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "人事档案管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "功能页面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Load 基本信息
End Sub

Private Sub Command2_Click()
Unload Me
Load 登录页面
End Sub

Private Sub Command3_Click()
Unload Me
Load 考勤信息
End Sub

Private Sub Command4_Click()
Unload Me
Load 条件查询
End Sub

Private Sub Command5_Click()
Unload Me
Load 调动信息
End Sub

Private Sub Command7_Click()
Unload Me
Load 登记页面
End Sub

Private Sub Command8_Click()
Unload Me
Load 删除页面
End Sub

Private Sub Command9_Click()
Unload Me
Load 修改页面
End Sub

Private Sub Form_Load()
Me.Show
End Sub
