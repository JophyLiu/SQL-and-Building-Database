VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ������ѯ 
   BackColor       =   &H00C0C0C0&
   Caption         =   "������ѯ"
   ClientHeight    =   6705
   ClientLeft      =   5310
   ClientTop       =   2040
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9825
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   5655
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2895
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         BackColorFixed  =   12632319
         BackColorSel    =   16777088
         BackColorBkg    =   12632256
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404040&
         Caption         =   "�˳�"
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
         Left            =   7320
         MaskColor       =   &H00404040&
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ѯ"
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
         Left            =   7320
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option2"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "��ѯ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ֵ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ֵ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����2"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   9840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ѯ������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "������ѯ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCnn As String             '���ݿ����Ӵ�
Public conODBC As Connection
Public Old_Row As Integer                 '�����Ϣ����ǰһ������
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
Dim value1_input, value2_input
Dim condition1_input, condition2_input
Dim log_input
'===�ж������Ƿ����
If Trim(Combo1.Text) = Trim(Combo2.Text) And Option1.Value = True Then
id = MsgBox("ͬһ����������ͬʱȡ����ֵ��������ѡ��", , "��ʾ")
Combo1.SetFocus
Exit Sub
End If
If Trim(Combo1.Text) = "" Or Trim(Text1.Text) = "" Then
id = MsgBox("����������1��������ֵ��", , "��ʾ")
Combo1.SetFocus
Exit Sub
End If
If Trim(Combo2.Text) <> "" And Trim(Text2.Text) = "" Then
id = MsgBox("����������2��ֵ��", , "��ʾ")
Text2.SetFocus
Exit Sub
End If
'===������������
condition1_input = Switch(Trim(Combo1.Text) = "Ա�����", "Staff_basic_info.staff_number", Trim(Combo1.Text) = "����", "staff_name", Trim(Combo1.Text) = "�Ա�", "staff_sex", Trim(Combo1.Text) = "����", "staff_where", Trim(Combo1.Text) = "ѧ��", "staff_ROFS", Trim(Combo1.Text) = "רҵ", "staff_major", Trim(Combo1.Text) = "ԭ����", "old_department", Trim(Combo1.Text) = "�²���", "new_department")
value1_input = Trim(Text1.Text)

condition2_input = Switch(Trim(Combo2.Text) = "Ա�����", "Staff_basic_info.staff_number", Trim(Combo2.Text) = "����", "staff_name", Trim(Combo2.Text) = "�Ա�", "staff_sex", Trim(Combo2.Text) = "����", "staff_where", Trim(Combo2.Text) = "ѧ��", "staff_ROFS", Trim(Combo2.Text) = "רҵ", "staff_major", Trim(Combo2.Text) = "ԭ����", "old_department", Trim(Combo2.Text) = "�²���", "new_department")
value2_input = Trim(Text2.Text)
If Option1.Value = True Then
log_input = "and"
Else
log_input = "or"
End If
'Sub_ConnectServer
'=====�����˲Ż�����Ϣ����¼
Dim strCnn As String
strCnn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RSXXGLXT;Data Source=PC-20151204RENG"
Dim conODBC As ADODB.Connection
Set conODBC = CreateObject("adodb.connection")
conODBC.Open strCnn

Dim cmd As ADODB.Command
Set cmd = CreateObject("adodb.command")
Dim rs As ADODB.Recordset
Set rs = CreateObject("adodb.recordset")
Dim param As ADODB.Parameter
Set param = CreateObject("adodb.parameter")

cmd.CommandType = CommandTypeEnum.adCmdText
If Trim(Combo2.Text) = "" Then
cmd.CommandText = "Select * From Staff_basic_info,Staff_attend_info,Staff_mobilize_info where Staff_basic_info.staff_number=Staff_attend_info.staff_number and Staff_attend_info.staff_number=Staff_mobilize_info.staff_number and " + condition1_input + " like ?"
Else
cmd.CommandText = "Select * From Staff_basic_info,Staff_attend_info,Staff_mobilize_info where Staff_basic_info.staff_number=Staff_attend_info.staff_number and Staff_attend_info.staff_number=Staff_mobilize_info.staff_number and (" + condition1_input + " like ? " + log_input + " " + condition2_input + " like ?)"
End If
'Set param = cmd.CreateParameter(condition1_input, adVarChar, adParamInput, 10)
cmd.Parameters.Append cmd.CreateParameter(condition1_input, adVarChar, adParamInput, 10)
'Set para = cmd.CreateParameter(condition2_input, adVarChar, adParamInput, 10)
If Trim(Combo2.Text) <> "" Then
cmd.Parameters.Append cmd.CreateParameter(condition2_input, adVarChar, adParamInput, 10)
End If
cmd.Parameters(0).Value = "%" + value1_input + "%"
If Trim(Combo2.Text) <> "" Then
cmd.Parameters(1).Value = "%" + value2_input + "%"
End If
cmd.ActiveConnection = conODBC
Set rs = cmd.Execute()
If rs.EOF And rs.BOF Then
id = MsgBox("û�в鵽��¼��", , "��ʾ")
Combo1.SetFocus
Exit Sub
Else
MSFlexGrid1.Rows = 1
While Not rs.EOF
MSFlexGrid1.AddItem rs!staff_number & Chr(9) & rs!staff_name & Chr(9) & rs!staff_sex & Chr(9) & rs!staff_where & Chr(9) & rs!staff_age & Chr(9) & rs!staff_birth & Chr(9) & rs!staff_add & Chr(9) & rs!staff_Email & Chr(9) & rs!staff_ROFS & Chr(9) & rs!staff_major & Chr(9) & rs!staff_intime & Chr(9) & rs!go_work_time & Chr(9) & rs!out_work_time & Chr(9) & rs!old_department & Chr(9) & rs!new_department    ''''''���ӱ��������
rs.MoveNext
Wend
End If
rs.Close
End Sub

Private Sub Command2_Click()
Unload Me
Load ����ҳ��
End Sub


Private Sub Form_Load()
Me.Show
MSFlexGrid1.FormatString = "^Ա����� |^  ����  |^ �Ա� |^   ����   |^ ����  |^   ����   |^  סַ  |^      Email      |^  ѧ��  |^ רҵ  |^���뱾��λʱ��|^�ϰ�ʱ��|^�°�ʱ��|^  ԭ����  |^  �²���  "
Combo1.List(0) = "Ա�����"
Combo1.List(1) = "����"
Combo1.List(2) = "�Ա�"
Combo1.List(3) = "����"
Combo1.List(4) = "ѧ��"
Combo1.List(5) = "רҵ"
Combo1.List(6) = "ԭ����"
Combo1.List(7) = "�²���"

Combo2.List(0) = ""
Combo2.List(1) = "Ա�����"
Combo2.List(2) = "����"
Combo2.List(3) = "�Ա�"
Combo2.List(4) = "����"
Combo2.List(5) = "ѧ��"
Combo2.List(6) = "רҵ"
Combo2.List(7) = "ԭ����"
Combo2.List(8) = "�²���"

End Sub
