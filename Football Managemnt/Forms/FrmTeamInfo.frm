VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmTeamInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Team Information"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsearchname 
      Height          =   390
      Left            =   2640
      TabIndex        =   18
      Top             =   240
      Width           =   3360
   End
   Begin VB.Frame Frame2 
      Caption         =   "Club Information"
      Height          =   1680
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   10455
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "FrmTeamInfo.frx":0000
         Height          =   885
         Left            =   90
         TabIndex        =   15
         Top             =   300
         Width           =   10230
         _ExtentX        =   18045
         _ExtentY        =   1561
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "MANAGER"
            Caption         =   "MANAGER"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "STADIUM"
            Caption         =   "STADIUM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "CAPACITY"
            Caption         =   "CAPACITY"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "FOUNDED"
            Caption         =   "FOUNDED"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "HONOURS"
            Caption         =   "HONOURS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2115.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2025.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3465.071
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8961
      TabIndex        =   8
      Top             =   8850
      Width           =   1425
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   9
      ToolTipText     =   "GO BACK TO CLUB SELECTION"
      Top             =   8700
      Width           =   1425
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   645
      TabIndex        =   4
      ToolTipText     =   "CLICK TO ADD PLAYER"
      Top             =   8850
      Width           =   1425
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2724
      TabIndex        =   5
      ToolTipText     =   "CLICK TO EDIT PLAYER INFO"
      Top             =   8850
      Width           =   1425
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4803
      TabIndex        =   6
      ToolTipText     =   "CLICK TO DELETE A PLAYER"
      Top             =   8850
      Width           =   1425
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6882
      TabIndex        =   7
      ToolTipText     =   "CLICK TO SAVE CHANGES"
      Top             =   8850
      Width           =   1425
   End
   Begin MSAdodcLib.Adodc TeamInfo 
      Height          =   330
      Left            =   3120
      Top             =   9360
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from players"
      Caption         =   "TeamInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   135
      TabIndex        =   1
      ToolTipText     =   "CLICK TO VIEW PICTURES"
      Top             =   3000
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   10054
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Team Profile"
      TabPicture(0)   =   "FrmTeamInfo.frx":0017
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Picture Gallery"
      TabPicture(1)   =   "FrmTeamInfo.frx":0033
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5235
         Left            =   90
         TabIndex        =   2
         Top             =   315
         Width           =   12105
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmTeamInfo.frx":004F
            Height          =   4560
            Left            =   135
            Negotiate       =   -1  'True
            TabIndex        =   3
            Top             =   390
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   8043
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   0   'False
            HeadLines       =   1
            RowHeight       =   17
            WrapCellPointer =   -1  'True
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Profiles"
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "NAME"
               Caption         =   "NAME"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "CLUB"
               Caption         =   "CLUB"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "POSITION"
               Caption         =   "POSITION"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "DOB"
               Caption         =   "DOB"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "COUNTRY"
               Caption         =   "COUNTRY"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "DOJ"
               Caption         =   "DOJ"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "TFROM"
               Caption         =   "TFROM"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "STATUS"
               Caption         =   "STATUS"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   2069.858
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1709.858
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   689.953
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5175
         Left            =   -74925
         TabIndex        =   10
         Top             =   450
         Width           =   11640
         Begin VB.TextBox txtthumbdata 
            Height          =   405
            Left            =   4935
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "for database"
            Top             =   2010
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.TextBox txtthumb 
            Height          =   465
            Left            =   4845
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "for loading the picture"
            Top             =   1170
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label lblstadname 
            Alignment       =   2  'Center
            Caption         =   "Home Ground"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   13
            Top             =   4635
            Width           =   8745
         End
         Begin VB.Label lblaway 
            Alignment       =   2  'Center
            Caption         =   "Away Kit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   780
            TabIndex        =   12
            Top             =   4320
            Width           =   990
         End
         Begin VB.Image ImageAway 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   1560
            Left            =   495
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   1560
         End
         Begin VB.Label lblhome 
            Alignment       =   2  'Center
            Caption         =   "Home Kit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   675
            TabIndex        =   11
            Top             =   2100
            Width           =   1125
         End
         Begin VB.Image ImageHome 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   1560
            Left            =   495
            Stretch         =   -1  'True
            Top             =   465
            Width           =   1560
         End
         Begin VB.Image ImageStad 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   4350
            Left            =   2595
            Stretch         =   -1  'True
            Top             =   210
            Width           =   8790
         End
      End
   End
   Begin MSAdodcLib.Adodc clubinfo 
      Height          =   345
      Left            =   5280
      Top             =   9360
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=football.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from clubs"
      Caption         =   "ClubInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Project1.desButton cmdSearch 
      Height          =   495
      Left            =   6240
      TabIndex        =   20
      Top             =   120
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
      caption         =   "Search"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Club Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   19
      Top             =   240
      Width           =   1440
   End
   Begin VB.Image ImgCntLogo 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1560
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Label lblteam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Team Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   11910
   End
End
Attribute VB_Name = "FrmTeamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim delete As Boolean
Dim confirm As Integer
Dim clubsql As String
'Dim rs As New ADODB.Recordset
Dim playersql As String

Private Sub cmdadd_Click()
    DataGrid1.AllowAddNew = True
    DataGrid1.AllowUpdate = True
    DataGrid1.Enabled = True
    
    cmdcan.Enabled = True
    cmdsave.Enabled = True
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
    
    DataGrid2.AllowAddNew = True
    DataGrid2.AllowUpdate = True
    DataGrid2.Enabled = True
End Sub

Private Sub cmdback_Click()
    FrmClub.Show
    Unload Me
End Sub

Private Sub cmdcan_Click()
    TeamInfo.RecordSource = playersql
    TeamInfo.Refresh
    clubinfo.RecordSource = clubsql
    clubinfo.Refresh
    delete = False
    
    DataGrid1.Refresh
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    DataGrid1.AllowDelete = False
    DataGrid1.Enabled = False
    
    DataGrid2.Refresh
    DataGrid2.AllowAddNew = False
    DataGrid2.AllowUpdate = False
    DataGrid2.AllowDelete = False
    DataGrid2.Enabled = False
    
    cmdcan.Enabled = False
    cmdsave.Enabled = False
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
End Sub

Private Sub cmddel_Click()
    Dim temp As Integer
    cmdcan.Enabled = True
    cmdsave.Enabled = True
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False

    DataGrid1.AllowUpdate = True
    DataGrid1.AllowDelete = True
    DataGrid1.Enabled = True
    delete = True
    temp = MsgBox("Select the record to be deleted and then press 'Save'", vbInformation, "To Delete")
End Sub

Private Sub cmdedit_Click()
    cmdcan.Enabled = True
    cmdsave.Enabled = True
    cmdadd.Enabled = False
    cmdedit.Enabled = False
    cmddel.Enabled = False
    
    DataGrid1.AllowUpdate = True
    DataGrid1.Enabled = True
    DataGrid2.AllowUpdate = True
    DataGrid2.Enabled = True
    
    ImageAway.Enabled = True
    ImageHome.Enabled = True
    ImageStad.Enabled = True
    ImgCntLogo.Enabled = True
End Sub

Private Sub cmdSave_Click()
    If delete = True Then
        confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
        If confirm = vbYes Then
            TeamInfo.Recordset.delete
            MsgBox "Record Deleted!", , "Message"
        Else
            MsgBox "Record Not Deleted!", , "Message"
        End If
    End If
    delete = False
    
    DataGrid1.Refresh
    DataGrid1.AllowAddNew = False
    DataGrid1.AllowUpdate = False
    DataGrid1.AllowDelete = False
    DataGrid1.Enabled = False
    
    DataGrid2.Refresh
    DataGrid2.AllowAddNew = False
    DataGrid2.AllowUpdate = False
    DataGrid2.AllowDelete = False
    DataGrid2.Enabled = False
    
    cmdcan.Enabled = False
    cmdsave.Enabled = False
    cmdadd.Enabled = True
    cmdedit.Enabled = True
    cmddel.Enabled = True
    
    ImageAway.Enabled = False
    ImageHome.Enabled = False
    ImageStad.Enabled = False
    ImgCntLogo.Enabled = False

End Sub

Private Sub cmdSearch_Click()

On Error GoTo NotFound

    Dim temp As String
    Dim temp2 As String
    Dim pic As String
               
        If Trim(txtsearchname.Text) = "" Then
            MsgBox "Please Enter The Appropriate Data", vbInformation, "Sys Man"
            txtsearchname.SetFocus
            SendKeys highLig
            Exit Sub
        End If

        TeamInfo.Refresh
        TeamInfo.Recordset.Find ("Club = '" & Trim(txtsearchname.Text) & "'")
        
        clubinfo.Refresh
        clubinfo.Recordset.Find ("Club_Name ='" & Trim(txtsearchname.Text) & "'")

    FrmTeamInfo.Caption = "Edit " & Trim(txtsearchname.Text) & " Information"
    
    clubsql = "select * from Clubs where Club_Name  like '" & Trim(txtsearchname.Text) & "'; "
    playersql = "select * from Players where Club like '" & Trim(txtsearchname.Text) & "';"
    
    TeamInfo.RecordSource = playersql
    TeamInfo.Refresh
    
    clubinfo.RecordSource = clubsql
    clubinfo.Refresh
    
    SSTab1.Tab = 0
    
    'lblstadname.Caption = clubinfo.Recordset.Fields("stadium").Value
    
    'txtthumb = clubinfo.Recordset.Fields("imgstad").Value & ""
    'ImageStad.Picture = LoadPicture(App.Path & txtthumb)
    
    'txtthumb = clubinfo.Recordset.Fields("imghome").Value & ""
    'ImageHome.Picture = LoadPicture(App.Path & txtthumb)
    
    'txtthumb = clubinfo.Recordset.Fields("imgaway").Value & ""
    'If txtthumb.Text <> "" Then
        'ImageAway.Picture = LoadPicture(App.Path & txtthumb)
    'End If
    
    'txtthumb = clubinfo.Recordset.Fields("logo").Value & ""
    'ImgCntLogo.Picture = LoadPicture(App.Path & txtthumb)
   
  Exit Sub

NotFound:
    MsgBox "The record you requested could not be found.", vbOKOnly + vbExclamation, "System Admin"
    txtsearchname.SetFocus
    SendKeys HiLyt
    
End Sub

Private Sub Form_Load()

    img = ""
    txtthumb.Text = ""
    txtthumbdata.Text = ""
    
    Call DataConn(TeamInfo, "Players")
    Call DataConn(clubinfo, "Clubs")
    
    cmdsave.Enabled = False
    cmdcan.Enabled = False
    
End Sub

Private Sub ImageAway_DblClick()
    img = "away"
    txtthumb.Text = ""
    frmBrowse.Show 1
    ImageAway.Picture = LoadPicture(txtthumb.Text)
    update
End Sub

Private Sub ImageHome_DblClick()
    img = "home"
    txtthumb.Text = ""
    frmBrowse.Show 1
    ImageHome.Picture = LoadPicture(txtthumb.Text)
    update
End Sub

Private Sub ImageStad_DblClick()
    img = "stadium"
    txtthumb.Text = ""
    frmBrowse.Show 1
    ImageStad.Picture = LoadPicture(txtthumb.Text)
    update
End Sub

Private Sub ImgCntLogo_DblClick()
    img = "logo"
    txtthumb.Text = ""
    frmBrowse.Show 1
    ImgCntLogo.Picture = LoadPicture(txtthumb.Text)
    update
End Sub

Private Sub update()
    Dim sqlstr As String
    sqlstr = "select * from clubs where club like '" & team & "';"
    
    rs.Open sqlstr, dbcon, adOpenStatic, adLockOptimistic
    
    Select Case img
        Case "home":
            rs.Fields("imghome") = "\images\kits\" & txtthumbdata.Text
        Case "away"
            rs.Fields("imgaway") = "\images\kits\" & txtthumbdata.Text
        Case "stadium"
            rs.Fields("imgstad") = "\images\stadiums\" & txtthumbdata.Text
        Case "logo"
            rs.Fields("logo") = "\images\logos\" & txtthumbdata.Text
    End Select
    rs.update
    rs.Close
End Sub

