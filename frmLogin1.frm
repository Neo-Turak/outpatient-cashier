VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin1 
   BorderStyle     =   0  'None
   Caption         =   "��¼"
   ClientHeight    =   4620
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   5625
   Icon            =   "frmLogin1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin1.frx":1082
   ScaleHeight     =   2729.648
   ScaleMode       =   0  'User
   ScaleWidth      =   5281.57
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin MSForms.Label Label6 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   1800
      Width           =   615
      ForeColor       =   16384
      VariousPropertyBits=   8388627
      Caption         =   "ID��"
      Size            =   "1085;661"
      FontName        =   "����"
      FontHeight      =   285
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label6 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   975
      ForeColor       =   16711680
      VariousPropertyBits=   8388627
      Caption         =   "���룺"
      Size            =   "1720;661"
      FontName        =   "����"
      FontHeight      =   285
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����շѹ���վ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   1864
      TabIndex        =   8
      Top             =   1080
      Width           =   1995
      VariousPropertyBits=   19
      Size            =   "3519;873"
      FontName        =   "����"
      FontHeight      =   180
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3047
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "�û���"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   1819
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      DataField       =   "����"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "ְλ"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin MSForms.TextBox TxtPassword 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
      VariousPropertyBits=   746604563
      BorderStyle     =   1
      Size            =   "4683;661"
      PasswordChar    =   42
      SpecialEffect   =   0
      FontName        =   "����"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextUserName 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1811
      Width           =   855
      VariousPropertyBits=   746604563
      BorderStyle     =   1
      Size            =   "1508;661"
      SpecialEffect   =   0
      FontName        =   "����"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SkinH_AttachEx Lib "D:\Users\NURA\vb 37��Ƥ��\SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'���ڽ�CreateRoundRectRgn������Բ�����򸳸�����
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'���ڴ���һ��Բ�Ǿ��Σ��þ�����X1��Y1-X2��Y2ȷ��������X3��Y3ȷ������Բ����Բ�ǻ��ȡ�
'���� ���ͼ�˵����
'X1,Y1 Long���������Ͻǵ�X��Y����
'X2,Y2 Long���������½ǵ�X��Y����
'X3 Long��Բ����Բ�Ŀ��䷶Χ��0��û��Բ�ǣ������ο�ȫԲ��
'Y3 Long��Բ����Բ�ĸߡ��䷶Χ��0��û��Բ�ǣ������θߣ�ȫԲ��
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��CreateRoundRectRgn����������ɾ�������Ǳ�Ҫ�ģ����򲻱�Ҫ��ռ�õ����ڴ�
Dim outrgn As Long
'����������һ��ȫ�ֱ���,�������������
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
If Label1.Caption = "����" Then
If Trim(Label2.Caption) = Trim(TxtPassword.Text) Then
    '�����ȷ������
        LoginSucceeded = True
        Me.Hide
          �շ�������.Show
       �շ�������.StatusBar1.Panels(3) = Label3.Caption
        �շ�������.StatusBar1.Panels(4) = Label4.Caption
       �շ�������.StatusBar1.Panels(5) = Label1.Caption
    Else
        MsgBox "��Ч��������û�����������!", , "��¼"
        TextUserName.SetFocus
        SendKeys "{Home}+{End}"
        End If
        Else
        MsgBox "�����ϵ��û����ͣ�"
        End If
End Sub

Private Sub CommandButton1_Click()
Dim x As Integer, y As Integer, Z As Integer
Z = (Me.Width - 4755) / 2
y = Me.Width / 2
x = Me.Height / 2   '�߶�
frmADODBLogon.Left = Me.Left + Z
frmADODBLogon.Top = VB.Screen.Height / 2 + x
frmADODBLogon.Show
End Sub

Private Sub Form_Activate() '����Activate()�¼�
Call rgnform(Me, 50, 50) '�����ӹ���
'SkinH_AttachEx "D:\Users\NURA\Desktop\���Ӳ���\Ƥ��\��Ө���.she", "" 'Ƥ������
End Sub
Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "�����Ѿ����У������ٴ�װ�أ�", vbOKOnly, "����"
 End
 End If
Dim x As Integer, y As Integer
x = Screen.Width / Screen.TwipsPerPixelX
y = Screen.Height / Screen.TwipsPerPixelY

End Sub

Private Sub Form_Unload(Cancel As Integer) '����Unload�¼�
DeleteObject outrgn '��Բ������ʹ�õ�����ϵͳ��Դ�ͷ�
End Sub
Private Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long) '�ӹ��̣��ı����fw��fh��ֵ��ʵ��Բ��
Dim w As Long, h As Long
w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels)
h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels)
outrgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub

Private Sub TextUserName_LostFocus()
Dim Conn As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Conn = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim ConnectString As String
ConnectString = "Provider=SQLOLEDB.1;password=sa;Persist Security Info=true;User ID=sa;Initial Catalog=ghgl;Data Source=TOP-PC"
Conn.Open ConnectString
Conn.CursorLocation = adUseClient
Mrc.Open "select * from �û��� where ID='" & TextUserName.Text & "'", Conn, adOpenKeyset, adLockOptimistic
    Set Label1.DataSource = Mrc
    Set Label2.DataSource = Mrc
    Set Label3.DataSource = Mrc
    Set Label2.DataSource = Mrc
    Set Label4.DataSource = Mrc
End Sub
