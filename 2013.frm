VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Ԥ������� 
   Caption         =   "Ԥ�������"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   825
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "2013.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "2013.frx":1082
   ScaleHeight     =   8389.496
   ScaleMode       =   0  'User
   ScaleWidth      =   5030.396
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text5 
      DataField       =   "����"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   20
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "סԺ��"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5160
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   19
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   18
      Text            =   "�ĵ���"
      Top             =   240
      Width           =   1212
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4320
      Top             =   4080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
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
      RecordSource    =   "סԺ��"
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "2013.frx":7DF3A
      Height          =   1695
      Left            =   0
      TabIndex        =   17
      Top             =   2880
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "���߱��"
         Caption         =   "���߱��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "����"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "�Ա�"
         Caption         =   "�Ա�"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "����"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "סԺ��"
         Caption         =   "סԺ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "סԺ��"
         Caption         =   "סԺ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "����"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "���"
         Caption         =   "���"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "����ҽ��"
         Caption         =   "����ҽ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "���֤��"
         Caption         =   "���֤��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "ҽ��֤��"
         Caption         =   "ҽ��֤��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "��ַ"
         Caption         =   "��ַ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "��Ժ����"
         Caption         =   "��Ժ����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "��������"
         Caption         =   "��������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "������"
         Caption         =   "������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "�տ�������"
         Caption         =   "�տ�������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "״̬"
         Caption         =   "״̬"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   461.801
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   740.412
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   253.022
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   284.829
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   379.772
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   398.665
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   253.022
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   278.372
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   442.908
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   474.476
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   455.583
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   272.154
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   499.826
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   474.476
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   461.801
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   582.094
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   316.397
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "���"
      Height          =   735
      Left            =   8040
      TabIndex        =   16
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton CMDExport 
      Caption         =   "����"
      Height          =   735
      Left            =   6120
      TabIndex        =   15
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   735
      Left            =   2160
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "200"
      Top             =   1440
      Width           =   972
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "Ԥ���"
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
   Begin VB.TextBox Text9 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-M-d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   12
      Top             =   840
      Width           =   1932
   End
   Begin VB.TextBox Text9 
      DataField       =   "��Ժ����"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   10
      Top             =   840
      Width           =   1932
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ˢ��"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11400
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "��ӡ"
      Height          =   735
      Left            =   4200
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "ҽ��֤��"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2280
      Width           =   2292
   End
   Begin VB.TextBox Text7 
      DataField       =   "���֤��"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      MaxLength       =   18
      TabIndex        =   7
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      DataField       =   "����"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   735
      Left            =   0
      TabIndex        =   26
      Top             =   2040
      Width           =   8535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "2013.frx":7DF4F
      Height          =   2055
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   14
      Top             =   4680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   13.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "��ˮ��"
         Caption         =   "��ˮ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "����"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "סԺ��"
         Caption         =   "סԺ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "����"
         Caption         =   "����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "���֤��"
         Caption         =   "���֤��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ҽ��֤��"
         Caption         =   "ҽ��֤��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "��Ժ����"
         Caption         =   "��Ժ����"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "��������"
         Caption         =   "��������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "������"
         Caption         =   "������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "�տ���"
         Caption         =   "�տ���"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   398.665
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   291.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   385.99
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   695.93
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   512.501
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   537.851
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   493.608
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   531.633
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   525.176
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   430.233
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   2880
      TabIndex        =   28
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   1080
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3600
      TabIndex        =   25
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8160
      TabIndex        =   24
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��    �ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ס Ժ �ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3720
      TabIndex        =   22
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8160
      TabIndex        =   21
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   8050.743
      X2              =   8050.743
      Y1              =   0
      Y2              =   2341.428
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "�������ڣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   11
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժ���ڣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   3135
   End
   Begin VB.Menu File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu print 
         Caption         =   "��ӡ��&P��"
         Shortcut        =   ^P
      End
      Begin VB.Menu save 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu export 
         Caption         =   "����(&O)"
      End
      Begin VB.Menu exit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu search 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu Clear 
         Caption         =   "���(&C)"
      End
   End
   Begin VB.Menu help 
      Caption         =   "����(&H)"
      Begin VB.Menu about 
         Caption         =   "����"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Ԥ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PRINTDLG Lib "comdlg32.dll" Alias "PrintDlgA" (pprintdlg As PRINTDLG) As Long
Private Type PRINTDLG
lStructSize As Long
hwndOwner As Long
hDevMode As Long
hDevNames As Long
hdc As Long
flags As Long
nFromPage As Integer
nToPage As Integer
nMinPage As Integer
nMaxpage As Integer
nCopies As Integer
hInstance As Integer
lCustDate As Long
lpfnPrintHook As Long
lpPrintTemplateName As String
lpSetupTemplateName As String
hPrintTempulate As Long
hSetupTempulate As Long
A As String
B As String
c As String
E As String
D As String
F As PictureBox
End Type

Private Sub about_Click()
frmAbout.Show
Me.Hide
End Sub

Private Sub cmdExport_Click()
   Dim i As Integer, r As Integer, c As Integer
   Dim newxls As Excel.Application
   Dim newbook As Excel.Workbook
   Dim newsheet As Excel.Worksheet
   Set newxls = CreateObject("Excel.Application") '����excelӦ�ó���,��excel2000
   Set newbook = newxls.Workbooks.Add '����������
   Set newsheet = newbook.Worksheets(1) '����������
   If SQL <> "" Then
   Adodc1.RecordSource = SQL
   Adodc1.Refresh
      End If
   If Adodc1.Recordset.RecordCount > 0 Then
   For i = 0 To DataGrid1.Columns.Count - 1
   newsheet.Cells(1, i + 1) = DataGrid1.Columns(i).Caption
   Next i   'ָ���������
   Adodc1.Recordset.MoveFirst
   Do Until Adodc1.Recordset.EOF
   r = Adodc1.Recordset.AbsolutePosition
   For c = 0 To DataGrid1.Columns.Count - 1
   DataGrid1.Col = c
   newsheet.Cells(r + 1, c + 1) = DataGrid1.Columns(c)
   Next c
   Adodc1.Recordset.MoveNext
   Loop
   
   Dim myval As Long
   Dim mystr As String
   myval = MsgBox("�Ƿ񱣴��Excel��?", vbYesNo, "��ʾ����")
   If myval = vbYes Then
   mystr = Date & Mid(Time, 1, 2)
   If Len(mystr) = 0 Then
   MsgBox "ϵͳ�������ļ�����Ϊ�գ�", , "��ʾ����"
   Exit Sub
   End If
   On Error GoTo ErrSave
   newsheet.SaveAs "C:\�κ����ݿ�\Excel����\" & mystr & ".xlsx"
   MsgBox "Excel�ļ�����ɹ���λ�ã�""C:\�κ����ݿ�\excel����\" & mystr & ".xlsx", , "��ʾ����"
   newxls.Quit
   Exit Sub
ErrSave:
   
   MsgBox Err.Description, , "��ʾ����"
   End If
   End If
   End Sub

Private Sub Command1_Click()
Call print_Click
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()

Dim DZ As String
Dim HZYLH As String
Dim DD As String
Dim CC As String
DD = Format(Date, "YYYMMDD")
HZYLH = Mid(Text8.Text, 3, 2) & Right(Text8.Text, 5)

If HZYLH = "" Then
MsgBox "����ҽ�ƺ���������ˮ�ŵǼǣ�����Ϊ�գ������ԣ�", vbInformation, "ע�����"
Else
CC = DD & HZYLH
Adodc2.Recordset.Fields("סԺ��") = Text1.Text
Adodc2.Recordset.Fields("��ַ") = Label6.Caption
Adodc2.Recordset.Fields("��Ժ����") = Trim(Text9(0).Text)
Adodc2.Recordset.Fields("��������") = Trim(Text9(1).Text)
Adodc2.Recordset.Fields("���֤��") = Text7.Text
Adodc2.Recordset.Fields("ҽ��֤��") = Text8.Text
Adodc2.Recordset.Fields("������") = Text2.Text
Adodc2.Recordset.Fields("�տ�������") = Label7.Caption
Adodc2.Recordset.Fields("״̬") = "�ѽ���"
Adodc2.Recordset.Update
With Adodc1.Recordset
.AddNew
.Fields("��ˮ��") = CC
.Fields("����") = Text5.Text
.Fields("סԺ��") = Text1.Text
.Fields("����") = Text4.Text
.Fields("���֤��") = Text7.Text
.Fields("ҽ��֤��") = Text8.Text
.Fields("��Ժ����") = Text9(0).Text
.Fields("��������") = Text9(1).Text
.Fields("������") = Text2.Text
.Fields("�տ���") = Label7.Caption
End With
Adodc1.Recordset.Update
Text1.Text = Val(Text1.Text) + 1
End If

On Error GoTo Er:    '������
Exit Sub
Er:
 MsgBox Err.Description, , "��ʾ����"
End Sub

Private Sub Command4_Click()
Unload Me
Load Me
End Sub

Private Sub exit_Click()
End
End Sub


Private Sub export_Click()
Call cmdExport_Click
End Sub

Private Sub Form_Load()
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=TOP-PC;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from סԺ�� where ״̬='���շ�'", Con, adOpenKeyset, adLockOptimistic
Set Adodc2.Recordset = Mrc
DataGrid2.Refresh

 Text9(1).Text = Date
 Label7.Caption = �շ�������.StatusBar1.Panels(3).Text
 Label8.Caption = �շ�������.StatusBar1.Panels(4).Text
 If Not Adodc1.Recordset.RecordCount = 0 Then
 Adodc1.Recordset.MoveLast
 Set Text1.DataSource = Adodc1
 Text1.Text = Val(Text1.Text) + 1
 Set Text1.DataSource = Nothing
 Else
 End If
 End Sub
Private Sub print_Click()
Printer.PaperSize = 9
'��ʼ��ӡ����
'���жԻ�����������
Printer.Orientation = 1
Printer.ScaleWidth = 20
Printer.ScaleHeight = 14
Printer.FontSize = 15
Printer.CurrentX = 2
Printer.CurrentY = 1.5
Printer.Print Text5.Text 'chuanghao
Printer.CurrentX = 9
Printer.CurrentY = 1.6
Printer.Print Text1.Text 'zhuyuanhao
Printer.CurrentX = 4
Printer.CurrentY = 2.1
Printer.Print Text4.Text  'xingming
Printer.CurrentX = 13
Printer.CurrentY = 2.15
Printer.Print Text8.Text 'yiliaozhenghao
Printer.CurrentX = 4
Printer.CurrentY = 2.7
Printer.Print Label6.Caption 'dizhi
Printer.CurrentX = 5
Printer.CurrentY = 3.25
Printer.Print Mid(Text9(0).Text, 1, 4)

Printer.CurrentX = 5.2
Printer.CurrentY = 3.9
Printer.Print Mid(Text9(0).Text, 6, 2)

Printer.CurrentX = 6.8
Printer.CurrentY = 3.9
Printer.Print Mid(Text9(0).Text, 9, 2)
Printer.CurrentX = 13
Printer.CurrentY = 3.25
Printer.Print Mid(Text9(1).Text, 1, 4)

Printer.CurrentX = 13.2
Printer.CurrentY = 3.9
Printer.Print Mid(Text9(1).Text, 6, 2)
Printer.CurrentX = 15
Printer.CurrentY = 3.9
Printer.Print Mid(Text9(1).Text, 9, 2)
Printer.CurrentX = 6
Printer.CurrentY = 4.5
Printer.Print Text2.Text + "Ԫ��"

Printer.CurrentX = 13
Printer.CurrentY = 4.3
Printer.Print "" '�տ���Ԥ��λ��
 If Len(Text2.Text) = 3 Then
 BAI = Mid(Text2.Text, 1, 1)
 SHI = Mid(Text2.Text, 2, 1)
 GE = Mid(Text2.Text, 3, 1)
  If Val(BAI) = 1 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
 End If
If Val(BAI) = 2 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 3 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 4 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 5 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 6 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "½"
End If
If Val(BAI) = 7 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 8 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 9 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "�N"
End If
If Val(BAI) = 0 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 1 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
End If
If Val(SHI) = 2 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 3 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 4 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 5 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 6 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "½"
End If
If Val(SHI) = 7 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 8 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 9 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "�N"
End If
If Val(SHI) = 0 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(GE) = 1 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
End If
If Val(GE) = 2 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 3 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 4 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 5 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 6 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "½"
End If
If Val(GE) = 7 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 8 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 9 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "�N"
End If
If Val(GE) = 0 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
End If
 If Len(Text2.Text) = 4 Then
 QIAN = Mid(Text2.Text, 1, 1)
 BAI = Mid(Text2.Text, 2, 1)
 SHI = Mid(Text2.Text, 3, 1)
 GE = Mid(Text2.Text, 4, 1)
 
 If Val(QIAN) = 1 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
End If
If Val(QIAN) = 2 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(QIAN) = 3 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(QIAN) = 4 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(QIAN) = 5 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(QIAN) = 6 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "½"
End If

If Val(QIAN) = 7 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(QIAN) = 8 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(QIAN) = 9 Then
Printer.CurrentX = 2
Printer.CurrentY = 5.1
Printer.Print "�N"
End If

 
 If Val(BAI) = 1 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
End If
If Val(BAI) = 2 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 3 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 4 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 5 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 6 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "½"
End If
If Val(BAI) = 7 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 8 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(BAI) = 9 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "�N"
End If
If Val(BAI) = 0 Then
Printer.CurrentX = 3.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(SHI) = 1 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
End If
If Val(SHI) = 2 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 3 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 4 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 5 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 6 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "½"
End If
If Val(SHI) = 7 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 8 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(SHI) = 9 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "�N"
End If
If Val(SHI) = 0 Then
Printer.CurrentX = 5
Printer.CurrentY = 5.1
Printer.Print "��"
End If

If Val(GE) = 1 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "Ҽ"
End If
If Val(GE) = 2 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 3 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 4 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 5 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 6 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "½"
End If
If Val(GE) = 7 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 8 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If
If Val(GE) = 9 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "�N"
End If
If Val(GE) = 0 Then
Printer.CurrentX = 6.5
Printer.CurrentY = 5.1
Printer.Print "��"
End If


End If
Printer.EndDoc
End Sub


Private Sub save_Click()
Call Command3_Click
End Sub

Private Sub search_Click()
Call Command4_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Text3_Change()
Label6.Caption = Text11.Text + Text6.Text + "��" + Text3.Text + "��"
If Len(Text3.Text) = 1 Then
Text2.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
End Sub

Private Sub Text8_Change()
If Text8.Text = "" Then Text6.Text = ""

If Not Text8.Text = "" Then
Text6.Text = Mid(Text8.Text, 3, 2)
End If
End Sub
