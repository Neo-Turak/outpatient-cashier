VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form �շ� 
   BackColor       =   &H8000000D&
   Caption         =   "�շ�"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   11760
   Begin VB.CommandButton Command3 
      Caption         =   "ȫ���շ�"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9600
      TabIndex        =   29
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      Height          =   405
      Left            =   8280
      TabIndex        =   28
      Text            =   "0"
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   405
      Left            =   5160
      TabIndex        =   26
      Text            =   "0"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   405
      Left            =   1320
      TabIndex        =   24
      Text            =   "0"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   22
      Text            =   "0"
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8280
      TabIndex        =   20
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   18
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Text            =   "0"
      Top             =   6840
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4200
      Top             =   9960
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=top-pc"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=top-pc"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "�շѼ�¼"
      Caption         =   "Adodc3"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "��鵥�շ�.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   14
      Top             =   8160
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2143
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "�շ�����"
         Caption         =   "�շ�����"
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
         DataField       =   "�շ�ʱ��"
         Caption         =   "�շ�ʱ��"
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
         DataField       =   "�շ�ԭ��"
         Caption         =   "�շ�ԭ��"
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
         DataField       =   "�շѽ��"
         Caption         =   "�շѽ��"
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
         DataField       =   "�շ���"
         Caption         =   "�շ���"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
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
      Height          =   450
      Left            =   1320
      TabIndex        =   12
      Text            =   "0"
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1320
      TabIndex        =   10
      Text            =   "0"
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�շ�"
      Height          =   495
      Left            =   9600
      TabIndex        =   9
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�շ�"
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
      Left            =   9720
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4680
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=top-pc"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=top-pc"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "���ﴦ��"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   8280
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
      Connect         =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=top-pc"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=ghgl;Data Source=top-pc"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "��鵥"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "��鵥�շ�.frx":0015
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2778
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   21
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
         DataField       =   "�����"
         Caption         =   "�����"
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
         DataField       =   "�����Ŀ"
         Caption         =   "�����Ŀ"
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
         DataField       =   "��λ"
         Caption         =   "��λ"
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
         DataField       =   "�۸�"
         Caption         =   "�۸�"
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
      BeginProperty Column09 
         DataField       =   "����ʱ��"
         Caption         =   "����ʱ��"
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
         DataField       =   "�������"
         Caption         =   "�������"
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
         DataField       =   "����ҽʦ"
         Caption         =   "����ҽʦ"
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
      BeginProperty Column13 
         DataField       =   "�����"
         Caption         =   "�����"
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
         DataField       =   "�������"
         Caption         =   "�������"
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
         DataField       =   "���ʱ��"
         Caption         =   "���ʱ��"
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
      BeginProperty Column17 
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
      BeginProperty Column18 
         DataField       =   "���ҽʦ"
         Caption         =   "���ҽʦ"
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
      BeginProperty Column19 
         DataField       =   "���ʱ��"
         Caption         =   "���ʱ��"
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
      BeginProperty Column20 
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "��鵥�շ�.frx":002A
      Height          =   2775
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4895
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   16
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "ҩƷ����"
         Caption         =   "ҩƷ����"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "�÷�"
         Caption         =   "�÷�"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
         DataField       =   "ҽ��"
         Caption         =   "ҽ��"
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
      BeginProperty Column15 
         DataField       =   "ʱ��"
         Caption         =   "ʱ��"
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
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "12345678901212346789"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "��ˮ�ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         DataField       =   "��������"
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
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         DataField       =   "�Ա�"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "4"
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
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "�շѽ�"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "���к���ҽ�Ʊ�����"
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "�ܽ�"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   120
      Top             =   1560
      Width           =   11415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "���к���ҽ�Ʊ�����"
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
      Left            =   2520
      TabIndex        =   21
      Top             =   6900
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "�Էѽ�"
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
      Left            =   6840
      TabIndex        =   19
      Top             =   4490
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "���к���ҽ�Ʊ�����"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "�Էѽ�"
      Height          =   405
      Left            =   6840
      TabIndex        =   13
      Top             =   6900
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   4480
      Width           =   975
   End
End
Attribute VB_Name = "�շ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Voice As SpVoice

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer
Dim c As Integer
Adodc2.Recordset.MoveFirst
i = Adodc2.Recordset.RecordCount
For c = 1 To i
Adodc2.Recordset.Fields("״̬") = "��ִ��"
Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF = True Then
Exit For
End If
Next c
Adodc2.Recordset.UpdateBatch
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields("��������") = Label2.Caption
Adodc3.Recordset.Fields("��ˮ��") = Text1.Text
Adodc3.Recordset.Fields("�շ�����") = Date
Adodc3.Recordset.Fields("�շ�ʱ��") = Time
Adodc3.Recordset.Fields("�շ�ԭ��") = "����ҩƷ��"
Adodc3.Recordset.Fields("�շѽ��") = Text2.Text
Adodc3.Recordset.Fields("�շ���") = �շ�������.StatusBar1.Panels(3).Text
Adodc3.Recordset.Update
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Integer
Dim c As Integer
Adodc1.Recordset.MoveFirst
i = Adodc1.Recordset.RecordCount
For c = 1 To i
Adodc1.Recordset.Fields("״̬") = "��ִ��"
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Exit For
End If
Next c
Adodc1.Recordset.UpdateBatch adAffectCurrent
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields("��������") = Label2.Caption
Adodc3.Recordset.Fields("��ˮ��") = Text1.Text
Adodc3.Recordset.Fields("�շ�����") = Date
Adodc3.Recordset.Fields("�շ�ʱ��") = Time
Adodc3.Recordset.Fields("�շ�ԭ��") = "��鵥��"
Adodc3.Recordset.Fields("�շѽ��") = Text3.Text
Adodc3.Recordset.Fields("�շ���") = �շ�������.StatusBar1.Panels(3).Text
End Sub
Private Sub Form_Activate()
On Error Resume Next
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=TOP-PC;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ���ﴦ�� where ��ˮ�� like'%" & Text1.Text & "%'and ״̬='���շ�'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
Set Adodc1.Recordset = Mrc
End Sub

Private Sub Form_Load()
Dim DD As String
DD = Format(Date, "YYYYMMDD")
Text1.Text = DD
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) = 15 Then
Dim Con As ADODB.Connection
Dim Mrc As ADODB.Recordset
Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Dim SQL As String
SQL = "Provider=sqloledb.1;Data Source=TOP-PC;Persist Security Info=true;user id=sa;password=sa;initial catalog=ghgl"
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ��鵥 where ��ˮ�� like'%" & Text1.Text & "%'and ״̬='���շ�'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid2.DataSource = Mrc
Set Adodc1.Recordset = Mrc
Set Label2.DataSource = Mrc
Set Label3.DataSource = Mrc
Set Label4.DataSource = Mrc

Set Con = New ADODB.Connection
Set Mrc = New ADODB.Recordset
Con.Open SQL
Con.CursorLocation = adUseClient
Mrc.Open "select * from ���ﴦ�� where ��ˮ�� like'%" & Text1.Text & "%'and ״̬='���շ�'", Con, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = Mrc
Set Adodc2.Recordset = Mrc

Dim i As Integer
Dim CC As Integer
i = Adodc2.Recordset.RecordCount
For CC = 1 To i
Text2.Text = Val(Text2.Text) + Adodc2.Recordset.Fields("���")
Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF = True Then
Exit For
End If
Next CC

Dim O As Integer
Dim VV As Integer
O = Adodc1.Recordset.RecordCount
For VV = 1 To O
Text3.Text = Val(Text3.Text) + Adodc1.Recordset.Fields("�۸�")
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Exit For
End If
Next VV
End If
End Sub

Private Sub Text10_Change()
Text10.Text = Val(Text6.Text) + Val(Text4.Text)
End Sub

Private Sub Text2_Change()
Text8.Text = Val(Text3.Text) + Val(Text2.Text)
End Sub

Private Sub Text3_Change()
Text8.Text = Val(Text3.Text) + Val(Text2.Text)

End Sub

Private Sub Text4_Change()
Text10.Text = Val(Text4) + Val(Text6.Text)
End Sub

Private Sub Text5_Change()
Text9.Text = Val(Text5.Text) + Val(Text7.Text)
End Sub

Private Sub Text6_Change()
Text6.Text = Val(Text2.Text) - Val(Text5.Text)
End Sub

Private Sub Text7_Change()
Text9.Text = Val(Text5.Text) + Val(Text7.Text)
End Sub
Private Sub Text9_LostFocus()
Text10.Text = Val(Text8.Text) - Val(Text9.Text)
End Sub
