VERSION 5.00
Begin VB.Form ����ҳ�� 
   BackColor       =   &H00C0C0C0&
   Caption         =   "����ѡ��"
   ClientHeight    =   5760
   ClientLeft      =   6180
   ClientTop       =   2175
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9255
   Begin VB.CommandButton Command6 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�޸���Ϣ"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "ɾ����Ϣ"
         Height          =   615
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "�Ǽ���Ϣ"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ѯ����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "������Ϣ"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "������ѯ"
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "������Ϣ"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "������Ϣ"
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
      Caption         =   "�汾   V 1.0"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���µ�������"
      BeginProperty Font 
         Name            =   "����"
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
Attribute VB_Name = "����ҳ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Load ������Ϣ
End Sub

Private Sub Command2_Click()
Unload Me
Load ��¼ҳ��
End Sub

Private Sub Command3_Click()
Unload Me
Load ������Ϣ
End Sub

Private Sub Command4_Click()
Unload Me
Load ������ѯ
End Sub

Private Sub Command5_Click()
Unload Me
Load ������Ϣ
End Sub

Private Sub Command7_Click()
Unload Me
Load �Ǽ�ҳ��
End Sub

Private Sub Command8_Click()
Unload Me
Load ɾ��ҳ��
End Sub

Private Sub Command9_Click()
Unload Me
Load �޸�ҳ��
End Sub

Private Sub Form_Load()
Me.Show
End Sub
