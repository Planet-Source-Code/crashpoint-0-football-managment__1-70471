VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewClub 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  New Club Registration"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoPlayer 
      Height          =   375
      Left            =   6360
      Top             =   9480
      Width           =   1440
      _ExtentX        =   2540
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
      Connect         =   $"frmNewClub.frx":0000
      OLEDBString     =   $"frmNewClub.frx":00A5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8640
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Club Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   3600
         ScaleHeight     =   1455
         ScaleWidth      =   1215
         TabIndex        =   46
         Top             =   1320
         Width           =   1215
         Begin VB.Image imgMan 
            Height          =   1455
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1215
         End
      End
      Begin Project1.desButton cmdPMAdd 
         Height          =   615
         Left            =   6240
         TabIndex        =   43
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         Caption         =   "Add Manager"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   975
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1720
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtPMLea 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtPMFound 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtPMClub 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtPMName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   1320
         Width           =   1935
      End
      Begin Project1.desButton cmdBrowseM 
         Height          =   615
         Left            =   4920
         TabIndex        =   42
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Caption         =   "Browse"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000009&
         Caption         =   "League"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
         Caption         =   "Founded"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   "Club Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         Caption         =   "Manager Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc PlayerInfo 
      Height          =   330
      Left            =   120
      Top             =   9480
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
      UserName        =   "admin"
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSAdodcLib.Adodc clubinfo 
      Height          =   330
      Left            =   1560
      Top             =   9480
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6180
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "CLICK TO VIEW PICTURES"
      Top             =   3120
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   10901
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   -2147483639
      TabCaption(0)   =   "&Team Profile"
      TabPicture(0)   =   "frmNewClub.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5175
         Left            =   -74925
         TabIndex        =   4
         Top             =   450
         Width           =   11640
         Begin VB.TextBox txtthumb 
            Height          =   465
            Left            =   4845
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "for loading the picture"
            Top             =   1170
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.TextBox txtthumbdata 
            Height          =   405
            Left            =   4935
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "for database"
            Top             =   2010
            Visible         =   0   'False
            Width           =   2520
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
         Begin VB.Image ImageHome 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   1560
            Left            =   495
            Stretch         =   -1  'True
            Top             =   465
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
            TabIndex        =   9
            Top             =   2100
            Width           =   1125
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
            TabIndex        =   8
            Top             =   4320
            Width           =   990
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
            TabIndex        =   7
            Top             =   4635
            Width           =   8745
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5595
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8025
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   3480
            ScaleHeight     =   1455
            ScaleWidth      =   1215
            TabIndex        =   47
            Top             =   4080
            Width           =   1215
            Begin VB.Image imgPlayer 
               Height          =   1455
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.TextBox txtPClub 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   38
            Top             =   4800
            Width           =   1935
         End
         Begin VB.TextBox txtPStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   29
            Top             =   3720
            Width           =   1935
         End
         Begin VB.TextBox txtPTFrom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   27
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txtPDoj 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   25
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtPCountry 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            TabIndex        =   23
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtPState 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   4440
            Width           =   1935
         End
         Begin VB.TextBox txtPDob 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Top             =   4080
            Width           =   1935
         End
         Begin VB.TextBox txtPPosition 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   3720
            Width           =   1935
         End
         Begin VB.TextBox txtPRegNo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox txtPId 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   13
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtPName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   2640
            Width           =   1935
         End
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmNewClub.frx":0166
            Height          =   2280
            Left            =   120
            Negotiate       =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   4022
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Profiles"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin Project1.desButton cmdBrowse 
            Height          =   735
            Left            =   4920
            TabIndex        =   41
            Top             =   4440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1296
            Caption         =   "Browse"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin Project1.desButton cmdPAddPlayer 
            Height          =   495
            Left            =   6480
            TabIndex        =   44
            Top             =   4200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "Add Player"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin Project1.desButton cmdPClear 
            Height          =   495
            Left            =   6480
            TabIndex        =   45
            Top             =   4800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label4 
            Caption         =   "Club Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   4800
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   28
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Transffer From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   26
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Date of Join"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   24
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Country"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   22
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "State"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Date of Birth"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Position"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Reg No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Id No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Player Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   2640
            Width           =   1095
         End
      End
   End
   Begin MSAdodcLib.Adodc adoFilter 
      Height          =   375
      Left            =   3240
      Top             =   9480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSAdodcLib.Adodc adoFilter2 
      Height          =   375
      Left            =   4920
      Top             =   9480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\football.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\football.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
End
Attribute VB_Name = "frmNewClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim picName As String
Dim spicname As String
Dim mpicname As String
Dim mspicname As String

Dim man As String
Dim pla As String

Private Sub cmdBrowse_Click()

    dlgCommon.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgCommon.ShowOpen
    
        picName = dlgCommon.FileName
        
            spicname = Mid$(picName, InStrRev(picName, "/") + 1)
        
            If picName <> "" Then
                imgPlayer.Picture = LoadPicture(picName)
                
            End If
            
End Sub

Private Sub cmdBrowseM_Click()

    dlgCommon.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgCommon.ShowOpen
    
        mpicname = dlgCommon.FileName
        
            mspicname = Mid$(mpicname, InStrRev(mpicname, "/") + 1)
        
            If mpicname <> "" Then
                imgMan.Picture = LoadPicture(mpicname)
                
            End If
            
End Sub

Public Sub BorGrid2()

On Error Resume Next
 'filter the books to show only the ones borrowed by the Current Borrower
    adoFilter2.RecordSource = "SELECT * FROM Clubs WHERE Club_Name = '" & Trim(txtPMClub.Text) & "'"
    adoFilter2.Refresh
     
    Set DataGrid2.DataSource = adoFilter2

    With DataGrid2
        .Columns(0).DataField = "Manager"
        .Columns(0).Caption = "Manager Name"
        .Columns(0).Width = 1500
        
        .Columns(1).DataField = "Club_Name"
        .Columns(1).Caption = "Club"
        .Columns(1).Width = 2000

        .Columns(2).DataField = "Founded"
        .Columns(2).Caption = "Founded"
        .Columns(2).Width = 800
        
        .Columns(5).DataField = "League"
        .Columns(5).Caption = "League"
        .Columns(5).Width = 1600
        
       
     End With
     
      'imgClub.Picture = LoadPicture("clubinfo.Recordset.Fields(9)")
                                    
End Sub
Public Sub BorGrid()

On Error Resume Next
 'filter the books to show only the ones borrowed by the Current Borrower
    adoFilter.RecordSource = "SELECT * FROM Players WHERE Club = '" & Trim(txtPClub.Text) & "'"
    adoFilter.Refresh
        
    Set DataGrid3.DataSource = adoFilter

    With DataGrid3
        .Columns(0).DataField = "Name"
        .Columns(0).Caption = "Name"
        .Columns(0).Width = 1500
        
        .Columns(1).DataField = "ID_No"
        .Columns(1).Caption = "Identification No"
        .Columns(1).Width = 2000

        .Columns(2).DataField = "Reg_No"
        .Columns(2).Caption = "Registration No"
        .Columns(2).Width = 2000
        
        .Columns(3).DataField = "Club"
        .Columns(3).Caption = "Club Name"
        .Columns(3).Width = 1300
        
        .Columns(4).DataField = "Position"
        .Columns(4).Caption = "Position"
        .Columns(4).Width = 800
        
        .Columns(5).DataField = "DOB"
        .Columns(5).Caption = "Date of Birth"
        .Columns(5).Width = 1500
        
        .Columns(6).DataField = "State"
        .Columns(6).Caption = "State"
        .Columns(6).Width = 1600
        
        .Columns(7).DataField = "Country"
        .Columns(7).Caption = "Country"
        .Columns(7).Width = 0
        
        .Columns(8).DataField = "DOJ"
        .Columns(8).Caption = "Date of Join"
        .Columns(8).Width = 1500
        
        .Columns(9).DataField = "TFrom"
        .Columns(9).Caption = "Transffered From"
        .Columns(9).Width = 1500
        
        .Columns(10).DataField = "Status"
        .Columns(10).Caption = "Status"
        .Columns(10).Width = 1500
        
        .Columns(11).DataField = "Yellow_Crd"
        .Columns(11).Caption = "Yellow Card"
        .Columns(11).Width = 0
        
        .Columns(12).DataField = "Red_Crd"
        .Columns(12).Caption = "Red Card"
        .Columns(12).Width = 0
        
     End With
     
      'imgClub.Picture = LoadPicture("clubinfo.Recordset.Fields(9)")
                                    
End Sub

Private Sub cmdPAddPlayer_Click()

'On Error GoTo noData 'handles expected unsuccessful data entry error
       
       'PlayerInfo.Refresh
       'PlayerInfo.Recordset.Find ("Reg_No = '" & Trim(txtPRegNo.Text) & "'")
       
       'PlayerInfo.Refresh
       'PlayerInfo.Recordset.Find ("Club = '" & Trim(txtPClub.Text) & "'")
       
        If Trim(txtPName.Text) = "" Or Trim(txtPId.Text) = "" Or Trim(txtPRegNo.Text) = "" _
            Or Trim(txtPPosition.Text) = "" Or Trim(txtPDob.Text) = "" _
            Or Trim(txtPState.Text) = "" Or Trim(txtPCountry.Text) = "" Or Trim(txtPDoj.Text) = "" _
            Or Trim(txtPtfrom.Text) = "" Or Trim(txtPStatus.Text) = "" Or Trim(txtPClub.Text) = "" Then
            
            Call missing
            
            'checks the missing field and focuses on it
            If Trim(txtPName.Text) = "" Then
                txtPName.Text = ""
                txtPName.SetFocus
            ElseIf Trim(txtPId.Text) = "" Then
                txtPId.Text = ""
                txtPId.SetFocus
            ElseIf Trim(txtPRegNo.Text) = "" Then
                txtPRegNo.Text = ""
                txtPRegNo.SetFocus
            ElseIf Trim(txtPClub.Text) = "" Then
                txtPClub.Text = ""
                txtPClub.SetFocus
            ElseIf Trim(txtPPosition.Text) = "" Then
                txtPPosition.Text = ""
                txtPPosition.SetFocus
            ElseIf Trim(txtPDob.Text) = "" Then
                txtPDob.Text = ""
                txtPDob.SetFocus
            ElseIf Trim(txtPState.Text) = "" Then
                txtPState.Text = ""
                txtPState.SetFocus
            ElseIf Trim(txtPCountry.Text) = "" Then
                txtPCountry.Text = ""
                txtPCountry.SetFocus
             ElseIf Trim(txtPDoj.Text) = "" Then
                txtPDoj.Text = ""
                txtPDoj.SetFocus
             ElseIf Trim(txtPtfrom.Text) = "" Then
                txtPtfrom.Text = ""
                txtPtfrom.SetFocus
             ElseIf Trim(txtPStatus.Text) = "" Then
                txtPStatus.Text = ""
                txtPStatus.SetFocus
            End If
        
            Exit Sub
        End If
        
        If IcValid(Trim(txtPId.Text)) = True Then
            MsgBox "The Player Id Is Already Exist " & vbCrLf & "Please Provide A Valid Id", vbInformation, "SysMan"
            txtPId.SetFocus
            SendKeys highLig
            Exit Sub
        End If
        
        'Call chkPlayer
        
        adoplayer.RecordSource = "SELECT * FROM Players WHERE Club = '" & Trim(txtPClub.Text) & "'"
        adoplayer.Refresh

        If adoplayer.Recordset.RecordCount > 24 Then
            MsgBox "The Team Only Can Occupied Maximum 25 Members" & vbCrLf & "You Have Reached The Maximum Entry", vbInformation, "FAS System"
            Exit Sub
        End If

        'PlayerInfo.Refresh
        'PlayerInfo.Recordset.Find ("Reg_No = '" & Trim(txtPRegNo.Text) & "'")

                PlayerInfo.Refresh
                PlayerInfo.Recordset.AddNew
            
            'On Error Resume Next
            
                With PlayerInfo.Recordset
                    .Fields(0) = txtPName.Text
                    .Fields(1) = txtPId.Text
                    .Fields(2) = txtPRegNo.Text
                    .Fields(3) = txtPClub.Text
                    .Fields(4) = txtPPosition.Text
                    .Fields(5) = txtPDob.Text
                    .Fields(6) = txtPState.Text
                    .Fields(7) = txtPDoj.Text
                    .Fields(8) = txtPtfrom.Text
                    .Fields(9) = txtPStatus.Text
                    .Fields(10) = "-"
                    .Fields(11) = "-"
                End With
                
                On Error Resume Next
                    If picName <> "" Then
                        PlayerInfo.Recordset.Fields(12) = picName
                    End If
                
                PlayerInfo.Recordset.Update
                PlayerInfo.Refresh
                PlayerInfo.Recordset.MoveFirst 'will generate an error if data has not been entered

                Call BorGrid
                 
                    txtPName.Text = ""
                    txtPId.Text = ""
                    txtPRegNo.Text = ""
                    txtPPosition.Text = ""
                    txtPDob.Text = ""
                    txtPState.Text = ""
                    txtPCountry.Text = ""
                    txtPDoj.Text = ""
                    txtPtfrom.Text = ""
                    txtPStatus.Text = ""
                    imgPlayer.Picture = LoadPicture("")
                    
                    PlayerInfo.Refresh
                   
                Exit Sub

'noData:
    'MsgBox "Player Identification Number already exists. Please Choose An Appropriate Registration Number", vbOKOnly, "System Admin"
    'txtPId.SetFocus
    'SendKeys highLig
    'Exit Sub
        
End Sub

Public Sub chkPlayer()

'On Error Resume Next
    
    adoplayer.RecordSource = "SELECT * FROM Players WHERE Club = '" & Trim(txtPClub.Text) & "'"
    adoplayer.Refresh

    If adoplayer.Recordset.RecordCount > 0 Then
        MsgBox "Invalid"
        Exit Sub
    End If
    
    Exit Sub
    
End Sub

Private Sub cmdPClear_Click()

    Unload frmNewClub
    
End Sub

Private Sub cmdPMAdd_Click()

   On Error GoTo ErrHandler 'handles expected unsuccessful data entry error

       'ClubInfo.Refresh
       'ClubInfo.Recordset.Find ("Club_Name = '" & Trim(txtPMClub.Text) & "'")

    'form field validation
    If Trim(txtPMName.Text) = "" Or Trim(txtPMClub.Text) = "" _
        Or Trim(txtPMFound.Text) = "" _
        Or Trim(txtPMLea.Text) = "" Then
        
        'MsgBox "Required field missing. Please fill up ALL the fields.", vbOKOnly + vbExclamation, "SysMan"
        Call missing
        
        'checks the missing field and focuses on it
        If Trim(txtPMName.Text) = "" Then
            txtPMName.Text = ""
            txtPMName.SetFocus
        ElseIf Trim(txtPMClub.Text) = "" Then
            txtPMClub.Text = ""
            txtPMClub.SetFocus
        End If
    
        Exit Sub
    End If

    'On Error Resume Next
    
    'man = ClubInfo.Recordset.Fields("Club_Name")
    
    'If Trim(txtPMClub.Text) <> man Then
                
        ClubInfo.Refresh
        ClubInfo.Recordset.AddNew
            
        'On Error Resume Next

            With ClubInfo.Recordset
                .Fields(0) = txtPMName.Text
                .Fields(1) = txtPMClub.Text
                .Fields(2) = txtPMFound.Text
                .Fields(3) = txtPMLea.Text
                .Fields(4) = mpicname
            End With
                    
            'If mpicname <> "" Then
                'ClubInfo.Recordset.Fields(4) = mpicname
            'End If
            
            txtPClub.Text = ClubInfo.Recordset.Fields(1)
            
            ClubInfo.Recordset.Update
            ClubInfo.Refresh
            ClubInfo.Recordset.MoveFirst 'will generate an error if data has not been entered
             
             Call BorGrid2
             
             txtPMName.Text = ""
             txtPMClub.Text = ""
             txtPMFound.Text = ""
             txtPMLea.Text = ""
                   
            ClubInfo.Refresh
               
            Exit Sub
           
    'Else
        'MsgBox "The Club Name is Already Exist." & vbCrLf + vbCrLf & "Please re-Name Your Club", "System Admin"
        'txtPMClub.Text = ""
        'txtPMClub.SetFocus
        'Exit Sub
        
   ' End If
'Exit Sub

ErrHandler:
    MsgBox "Club name already exists. Please choose a different Club name.", vbOKOnly, "System Admin"
    txtPMClub.SetFocus
    txtPClub.Text = ""
    SendKeys highLig
    'Exit Sub

            
End Sub


Private Sub Form_Load()

    Call DataConn(PlayerInfo, "Players")
    Call DataConn(ClubInfo, "Clubs")
        
    adoFilter.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
        
     'Sets the command type to Table
        adoFilter.CommandType = adCmdText
        
     adoFilter2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
        
     'Sets the command type to Table
        adoFilter2.CommandType = adCmdText
        
    PlayerInfo.Recordset.MoveFirst
    ClubInfo.Recordset.MoveFirst

End Sub

