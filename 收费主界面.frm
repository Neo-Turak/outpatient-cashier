VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm �շ������� 
   BackColor       =   &H8000000C&
   Caption         =   "�ۺ��շѹ���վ"
   ClientHeight    =   9210
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11610
   LinkTopic       =   "MDIForm1"
   Picture         =   "�շ�������.frx":0000
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8835
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   3016
            TextSave        =   "2016-06-10"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "17:12"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "ʱ��"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "��ǰ�û�"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "���շ�����"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6641
            MinWidth        =   5292
            Text            =   "�ĵ�������Ժ"
            TextSave        =   "�ĵ�������Ժ"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "ҽԺ����"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "��������"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu �շ���Ŀ 
      Caption         =   "�շ���Ŀ(&D)"
      Begin VB.Menu �����շ� 
         Caption         =   "�����շ�"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Ԥ������� 
         Caption         =   "Ԥ�������"
         Shortcut        =   {F4}
      End
      Begin VB.Menu ҹ���շѹ��� 
         Caption         =   "ҹ���շѹ���"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu ���Ի����� 
      Caption         =   "���Ի�����(&G)"
      Begin VB.Menu �޸����� 
         Caption         =   "�޸�����"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ��Ļ���� 
         Caption         =   "��Ļ����"
      End
      Begin VB.Menu �˳�ϵͳ 
         Caption         =   "�˳�ϵͳ"
      End
   End
End
Attribute VB_Name = "�շ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ��鵥�շ�_Click()

End Sub

Private Sub �����շ�_Click()
�շ�.Show
End Sub

Private Sub �޸�����_Click()
�����޸�.Show
End Sub

Private Sub ҹ���շѹ���_Click()
סԺ�շѼ�¼.Show
End Sub

Private Sub Ԥ�������_Click()
�շѹ���վ.Ԥ�������.Show
End Sub
