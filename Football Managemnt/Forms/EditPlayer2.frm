VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EditPlayer2 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :. Edit Player Information"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   6720
      Width           =   7215
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   5640
         TabIndex        =   28
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
      Begin Project1.desButton cmdUpdate 
         Height          =   375
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Update"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.desButton cmdReload 
         Height          =   375
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Reload"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.TextBox txtsearchname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   26
      Top             =   300
      Width           =   3360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Player Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7200
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   5760
         ScaleHeight     =   1455
         ScaleWidth      =   1215
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
         Begin VB.Image imageplayer 
            Height          =   1485
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.TextBox txtRedCrd 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4005
         Width           =   2370
      End
      Begin VB.TextBox txtPYelCrd 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   4725
         Width           =   2370
      End
      Begin VB.TextBox txtPStatus 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3285
         Width           =   2370
      End
      Begin VB.TextBox txtPState 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2520
         Width           =   2370
      End
      Begin VB.TextBox txtPReg 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1845
         Width           =   2370
      End
      Begin VB.TextBox txtPId 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   7
         Top             =   1845
         Width           =   2370
      End
      Begin VB.TextBox txtPClub 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1155
         Width           =   2370
      End
      Begin VB.TextBox txtPDoj 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3255
         Width           =   2370
      End
      Begin VB.TextBox txtPPos 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1155
         Width           =   2370
      End
      Begin VB.TextBox txtPDob 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2520
         Width           =   2370
      End
      Begin VB.TextBox txtPtfrom 
         BackColor       =   &H8000000A&
         DataSource      =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3975
         Width           =   2370
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin Project1.desButton cmdBrowse 
         Height          =   735
         Left            =   5880
         TabIndex        =   14
         Top             =   2760
         Width           =   975
         _ExtentX        =   1720
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Player  Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "Red Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   25
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "Yellow Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   23
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Registration No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   21
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Identification No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblclub 
         BackColor       =   &H80000009&
         Caption         =   "Club Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label lbldoj 
         BackColor       =   &H80000009&
         Caption         =   "Date of Join"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   2955
         Width           =   1455
      End
      Begin VB.Label lblposition 
         BackColor       =   &H80000009&
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   17
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label lbldob 
         BackColor       =   &H80000009&
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lbltfrom 
         BackColor       =   &H80000009&
         Caption         =   "Transferred from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   3675
         Width           =   1800
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6000
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.desButton cmdSearch 
      Height          =   495
      Left            =   5520
      TabIndex        =   31
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Search"
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
   Begin MSAdodcLib.Adodc Player 
      Height          =   390
      Left            =   1440
      Top             =   7800
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   688
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
      Caption         =   "Search"
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
   Begin MSAdodcLib.Adodc Club 
      Height          =   390
      Left            =   3720
      Top             =   7800
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   688
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
      Caption         =   "Search Club"
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
   Begin Project1.desButton cmdNewSearch 
      Height          =   495
      Left            =   600
      TabIndex        =   32
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "New Search"
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
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Player Id : -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   33
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "EditPlayer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mpicname As String
Dim mspicname As String

Private Sub cmdBrowse_Click()

    dlgCommon.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
    dlgCommon.ShowOpen
    
        mpicname = dlgCommon.FileName
        
            mspicname = Mid$(mpicname, InStrRev(mpicname, "/") + 1)
        
            If mpicname <> "" Then
                imageplayer.Picture = LoadPicture(mpicname)
                
            End If

End Sub

Private Sub cmdClose_Click()

    Unload EditPlayer2
    
End Sub

Private Sub cmdNewSearch_Click()
        
        txtsearchname.Text = ""
        txtName.Text = ""
        txtPId.Text = ""
        txtPReg.Text = ""
        txtPClub.Text = ""
        txtPPos.Text = ""
        txtPDob.Text = ""
        txtPState.Text = ""
        txtPDoj.Text = ""
        txtPtfrom.Text = ""
        txtPStatus.Text = ""
        txtPYelCrd.Text = ""
        txtRedCrd.Text = ""
        imageplayer.Picture = LoadPicture("")
        
        cmdReload.Enabled = False
        cmdUpdate.Enabled = False
        
        txtName.Enabled = False
        txtPId.Enabled = False
        txtPReg.Enabled = False
        txtPClub.Enabled = False
        txtPPos.Enabled = False
        txtPDob.Enabled = False
        txtPState.Enabled = False
        txtPDoj.Enabled = False
        txtPtfrom.Enabled = False
        txtPStatus.Enabled = False
        txtPYelCrd.Enabled = False
        txtRedCrd.Enabled = False
        
        txtName.Locked = True
        txtPId.Locked = True
        txtPReg.Locked = True
        txtPClub.Locked = True
        txtPPos.Locked = True
        txtPDob.Locked = True
        txtPState.Locked = True
        txtPDoj.Locked = True
        txtPtfrom.Locked = True
        txtPStatus.Locked = True
        txtPYelCrd.Locked = True
        txtRedCrd.Locked = True
        
End Sub

Private Sub cmdReload_Click()

    If MsgBox("This will reload the current Player's profile !!!." & vbCrLf & _
            " Any unsaved data will be lost. Do You Want To Proceed?", _
                            vbYesNo + vbQuestion, "SysMan") = vbYes Then
        
        imageplayer.Picture = LoadPicture("")
        dlgCommon.FileName = ""
        
        Player.Refresh
        Player.Recordset.Find ("ID_No = '" & Trim(txtsearchname.Text) & "'")
         
        On Error Resume Next
        
        txtName.Text = Player.Recordset.Fields("Name")
        txtPId.Text = Player.Recordset.Fields("ID_No")
        txtPReg.Text = Player.Recordset.Fields("Reg_No")
        txtPClub.Text = Player.Recordset.Fields("Club")
        txtPPos.Text = Player.Recordset.Fields("Position")
        txtPDob.Text = Player.Recordset.Fields("DOB")
        txtPState.Text = Player.Recordset.Fields("State")
        txtPDoj.Text = Player.Recordset.Fields("DOJ")
        txtPtfrom.Text = Player.Recordset.Fields("TFrom")
        txtPStatus.Text = Player.Recordset.Fields("Status")
        txtPYelCrd.Text = Player.Recordset.Fields("Yellow_Crd")
        txtRedCrd.Text = Player.Recordset.Fields("Red_Crd")
        
        On Error Resume Next
            imageplayer.Picture = LoadPicture("")
            
        txtName.Enabled = True
        txtPId.Enabled = True
        txtPReg.Enabled = True
        txtPClub.Enabled = False
        txtPPos.Enabled = True
        txtPDob.Enabled = True
        txtPState.Enabled = True
        txtPDoj.Enabled = True
        txtPtfrom.Enabled = True
        txtPStatus.Enabled = True
        txtPYelCrd.Enabled = True
        txtRedCrd.Enabled = True
        
        txtName.Locked = False
        txtPId.Locked = False
        txtPReg.Locked = False
        txtPClub.Locked = True
        txtPPos.Locked = False
        txtPDob.Locked = False
        txtPState.Locked = False
        txtPDoj.Locked = False
        txtPtfrom.Locked = False
        txtPStatus.Locked = False
        txtPYelCrd.Locked = False
        txtRedCrd.Locked = False
                 
        txtName.BackColor = &HFFFFFF
        txtPId.BackColor = &HFFFFFF
        txtPReg.BackColor = &HFFFFFF
        txtPClub.BackColor = &HFFFFFF
        txtPPos.BackColor = &HFFFFFF
        txtPDob.BackColor = &HFFFFFF
        txtPState.BackColor = &HFFFFFF
        txtPDoj.BackColor = &HFFFFFF
        txtPtfrom.BackColor = &HFFFFFF
        txtPStatus.BackColor = &HFFFFFF
        txtPYelCrd.BackColor = &HFFFFFF
        txtRedCrd.BackColor = &HFFFFFF
                                         
            cmdBrowse.Enabled = True
            
            cmdUpdate.Enabled = True
            cmdReload.Enabled = True
            
            txtPClub.SetFocus
            SendKeys highLig
    Else
        Exit Sub
    
    End If

End Sub

Private Sub cmdSearch_Click()

Dim pic As String
Dim CPic As String
    
On Error GoTo NotFound
    
    If txtsearchname.Text = "" Then
        MsgBox "Please Enter Appropriate Value", vbCritical, "System Admin"
        txtsearchname.SetFocus
        SendKeys highLig
        Exit Sub
    End If
    
        Player.Refresh
        Player.Recordset.Find ("ID_No = '" & Trim(txtsearchname.Text) & "'")
         
        On Error Resume Next
        
        txtName.Text = Player.Recordset.Fields("Name")
        txtPId.Text = Player.Recordset.Fields("ID_No")
        txtPReg.Text = Player.Recordset.Fields("Reg_No")
        txtPClub.Text = Player.Recordset.Fields("Club")
        txtPPos.Text = Player.Recordset.Fields("Position")
        txtPDob.Text = Player.Recordset.Fields("DOB")
        txtPState.Text = Player.Recordset.Fields("State")
        txtPDoj.Text = Player.Recordset.Fields("DOJ")
        txtPtfrom.Text = Player.Recordset.Fields("TFrom")
        txtPStatus.Text = Player.Recordset.Fields("Status")
        txtPYelCrd.Text = Player.Recordset.Fields("Yellow_Crd")
        txtRedCrd.Text = Player.Recordset.Fields("Red_Crd")
        pic = Player.Recordset.Fields("Picture")
            
            On Error Resume Next
            imageplayer.Picture = LoadPicture(pic)
        
        txtsearchname.SetFocus
        SendKeys HiLyt
        
        cmdReload.Enabled = True
        
        Exit Sub
NotFound:
        MsgBox "The player profile you requested could not be found.", vbOKOnly + vbExclamation, "System Admin"
        
        imageplayer.Picture = LoadPicture("")
        txtsearchname.SetFocus
        SendKeys HiLyt

End Sub

Private Sub cmdUpdate_Click()

On Error GoTo errorhandle
    
    If txtPClub.Text = "" Then
        Call missing
        txtPClub.SetFocus
        Exit Sub
    End If
    
    If txtPId.Text = "" Then
        Call missing
        txtPId.SetFocus
        Exit Sub
    End If
        
    If txtPReg.Text = "" Then
        Call missing
        txtPReg.SetFocus
        Exit Sub
    End If
    
    If txtPPos.Text = "" Then
        Call missing
        txtPPos.SetFocus
        Exit Sub
    End If
    
    If txtPDob.Text = "" Then
        Call missing
        txtPDob.SetFocus
        Exit Sub
    End If
    
    If txtPState.Text = "" Then
        Call missing
        txtPState.SetFocus
        Exit Sub
    End If
    
    If txtPDoj.Text = "" Then
        Call missing
        txtPDoj.SetFocus
        Exit Sub
    End If
    
    If txtPtfrom.Text = "" Then
        Call missing
        txtPtfrom.SetFocus
        Exit Sub
    End If
    
    If txtPStatus.Text = "" Then
        Call missing
        txtPStatus.SetFocus
        Exit Sub
    End If
    
    If txtPYelCrd.Text = "" Then
        Call missing
        txtPYelCrd.SetFocus
        Exit Sub
    End If
    
    If txtRedCrd.Text = "" Then
        Call missing
        txtRedCrd.SetFocus
        Exit Sub
    End If
        
    'If txtPId.Text = Search.Recordset.Fields("ID_No") Then
        'MsgBox "Player Registration Number Already Exist !!!", vbInformation, "SysMan"
        'txtPId.Text = ""
        'txtPId.SetFocus
        'Exit Sub
    'End If

    If MsgBox("The Player Profile Will Be Updated.." & vbCrLf & _
            " Proceed ?", vbOKCancel + vbQuestion, "SysMan") = vbOK Then
        
        Player.Recordset.Fields("Name") = txtName.Text
        Player.Recordset.Fields("ID_No") = txtPId.Text
        Player.Recordset.Fields("Reg_No") = txtPReg.Text
        Player.Recordset.Fields("Club") = txtPClub.Text
        Player.Recordset.Fields("Position") = txtPPos.Text
        Player.Recordset.Fields("DOB") = txtPDob.Text
        Player.Recordset.Fields("State") = txtPState.Text
        Player.Recordset.Fields("DOJ") = txtPDoj.Text
        Player.Recordset.Fields("TFrom") = txtPtfrom.Text
        Player.Recordset.Fields("Status") = txtPStatus.Text
        Player.Recordset.Fields("Yellow_Crd") = txtPYelCrd.Text
        Player.Recordset.Fields("Red_Crd") = txtRedCrd.Text
             
        On Error Resume Next
        If dlgCommon.FileName <> "" Then
            Player.Recordset.Fields("Picture") = dlgCommon.FileName
        End If
        
        Player.Recordset.Update
        Player.Refresh
            
        If MsgBox("Record successfully updated. Continue editing records?", vbYesNo + vbQuestion, "System Manager") = vbYes Then
            txtsearchname.SetFocus
            SendKeys HiLyt
            Exit Sub
        Else
            Unload Me
        End If
    Else
            Exit Sub

    End If
    
    Exit Sub
    
errorhandle:
End Sub

Private Sub Form_Load()

    Call DataConn(Player, "Players")
    Call DataConn(Club, "Clubs")

End Sub
