VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm 收费主界面 
   BackColor       =   &H8000000C&
   Caption         =   "综合收费工作站"
   ClientHeight    =   9210
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11610
   LinkTopic       =   "MDIForm1"
   Picture         =   "收费主界面.frx":0000
   StartUpPosition =   2  '屏幕中心
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
            Object.ToolTipText     =   "日期"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "17:12"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "时间"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "当前用户"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "科室"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "待收费数量"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6641
            MinWidth        =   5292
            Text            =   "荒地镇卫生院"
            TextSave        =   "荒地镇卫生院"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "医院名称"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "华文中宋"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu 收费项目 
      Caption         =   "收费项目(&D)"
      Begin VB.Menu 常规收费 
         Caption         =   "常规收费"
         Shortcut        =   {F2}
      End
      Begin VB.Menu 预交款管理 
         Caption         =   "预交款管理"
         Shortcut        =   {F4}
      End
      Begin VB.Menu 夜班收费管理 
         Caption         =   "夜班收费管理"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu 个性化设置 
      Caption         =   "个性化设置(&G)"
      Begin VB.Menu 修改密码 
         Caption         =   "修改密码"
      End
   End
   Begin VB.Menu 其他 
      Caption         =   "其他"
      Begin VB.Menu 屏幕保护 
         Caption         =   "屏幕保护"
      End
      Begin VB.Menu 退出系统 
         Caption         =   "退出系统"
      End
   End
End
Attribute VB_Name = "收费主界面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub 检查单收费_Click()

End Sub

Private Sub 常规收费_Click()
收费.Show
End Sub

Private Sub 修改密码_Click()
密码修改.Show
End Sub

Private Sub 夜班收费管理_Click()
住院收费记录.Show
End Sub

Private Sub 预交款管理_Click()
收费工作站.预交款管理.Show
End Sub
