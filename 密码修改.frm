VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form �����޸� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����޸�"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "�����޸�.frx":0000
   ScaleHeight     =   4245
   ScaleWidth      =   5550
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "�û���"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      DataField       =   "����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
      VariousPropertyBits=   19
      Caption         =   "ȷ  ��"
      Size            =   "2566;873"
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TextBox4 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3080
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      PasswordChar    =   42
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2480
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      PasswordChar    =   42
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox2 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1900
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      PasswordChar    =   42
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
      VariousPropertyBits=   746604563
      Size            =   "4048;661"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   3120
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "�����룺"
      Size            =   "1720;582"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "�����룺"
      Size            =   "1720;582"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   1960
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "ԭ���룺"
      Size            =   "1720;582"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1350
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "�û�����"
      Size            =   "1720;582"
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "�����޸�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If TextBox2.Text = Label2.Caption Then

   If TextBox3.Text = TextBox4.Text Then
         Label2.Caption = TextBox3.Text
         Adodc1.Recordset.UpdateBatch adAffectCurrent
         
   MsgBox "�޸�����ɹ������µ�¼��"
   End
   Else
   MsgBox "��������Ч������������", vbInformation, "����"
   TextBox3.Text = ""
   TextBox4.Text = ""
   TextBox3.SetFocus
   End If

Else
MsgBox " ԭ�������", vbInformation, "�������"
TextBox2.SetFocus
End If
End Sub

Private Sub Form_Activate()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=TOP-PC;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from �û��� where �û���='" & TextBox1.Text & "'and ����='" & Label3.Caption & "'and ְλ='" & Label4.Caption & "'", Con, adOpenKeyset, adLockOptimistic
Set Adodc1.Recordset = Mrc
End Sub

Private Sub Form_Initialize()
Me.Width = 5685
Me.Height = 4695
End Sub

Private Sub Form_Load()
TextBox1.Text = �շ�������.StatusBar1.Panels(3).Text
Label3.Caption = �շ�������.StatusBar1.Panels(4).Text
Label4.Caption = �շ�������.StatusBar1.Panels(5).Text
End Sub


