VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ��¼ҳ�� 
   BackColor       =   &H00C0C0C0&
   Caption         =   "������Ϣ����ϵͳ"
   ClientHeight    =   6930
   ClientLeft      =   4875
   ClientTop       =   1890
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   11460
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "conODBC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��¼���"
      Height          =   3015
      Left            =   2760
      TabIndex        =   4
      Top             =   3000
      Width           =   5415
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000015&
         Caption         =   "�˳�"
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000005&
         Caption         =   "��¼"
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "�û���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "2015��12��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "����   ����ũҵ��ѧ��ѧ����ϢѧԺ   13�ſ�1��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   10455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "�汾   V 1.0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "������Ϣ����ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1200
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   10500
   End
End
Attribute VB_Name = "��¼ҳ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCnn As String             '���ݿ����Ӵ�
Public conODBC As Connection
'���ݿ����Ӷ���
'-------�������ݿ������
Public Sub Sub_ConnectServer()
Dim strCnn As String
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Dim conODBC As ADODB.Connection
Set conODBC = CreateObject("adodb.connection")
conODBC.Open strCnn
End Sub
Private Sub Command1_Click()
Dim name_user
Dim password_user
If Text1.Text = "" Then
id = MsgBox("�������û�����", , "��ʾ")
Exit Sub
'Else
'name_user = Trim("Text1.Text")
End If
If Text2.Text = "" Then
id = MsgBox("��������", , "��ʾ")
Exit Sub
'Else
'passsword_user = Trim(Text2.Text)
End If
'=========================��ADO��ʽ�����ݿ�
'MsgBox"Opening rcgl_sys...�˲Ź������ݿ�"
'Sub_ConnectServer
Dim strCnn As String
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Dim conODBC As ADODB.Connection
Set conODBC = CreateObject("adodb.connection")
conODBC.Open strCnn
 
Dim rs As ADODB.Recordset
Set rs = CreateObject("adodb.recordset")
rs.CursorType = adOpenStatic
rs.LockType = adLockOptimistic
'---------------��ϵͳ�û���

rs.CursorLocation = adUseClient

rs.Open "Select * From login where username='" + Text1.Text + "'", conODBC, adOpenStatic, adLockReadOnly
num_records = rs.RecordCount
If num_records = 0 Then
id = MsgBox("�û�������ȷ�����������룡", , "")
rs.Close
Text1.SetFocus
Exit Sub
Else
If Trim(rs!passkey) <> Trim(Text2.Text) Then
id = MsgBox("�����ȷ�����������룡", , "")
rs.Close
Text2.SetFocus
Exit Sub
End If
End If
rs.Close
Unload Me
Load ����ҳ��
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Show
Text1.SetFocus
End Sub

