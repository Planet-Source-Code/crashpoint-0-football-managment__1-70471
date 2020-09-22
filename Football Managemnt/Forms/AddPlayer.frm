VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AddPlayer 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :. Add New Player Information"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoplayer 
      Height          =   375
      Left            =   4440
      Top             =   6120
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
      TabIndex        =   4
      Top             =   120
      Width           =   7200
      Begin VB.TextBox txtName 
         BackColor       =   &H80000009&
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
         TabIndex        =   17
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtPtfrom 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   16
         Top             =   3975
         Width           =   2370
      End
      Begin VB.TextBox txtPDob 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   15
         Top             =   2520
         Width           =   2370
      End
      Begin VB.TextBox txtPPos 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   14
         Top             =   1155
         Width           =   2370
      End
      Begin VB.TextBox txtPDoj 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   13
         Top             =   3255
         Width           =   2370
      End
      Begin VB.TextBox txtPClub 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   12
         Top             =   1155
         Width           =   2370
      End
      Begin VB.TextBox txtPId 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1845
         Width           =   2370
      End
      Begin VB.TextBox txtPReg 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   10
         Top             =   1845
         Width           =   2370
      End
      Begin VB.TextBox txtPState 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   9
         Top             =   2520
         Width           =   2370
      End
      Begin VB.TextBox txtPStatus 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   8
         Top             =   3285
         Width           =   2370
      End
      Begin VB.TextBox txtPYelCrd 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   7
         Top             =   4725
         Width           =   2370
      End
      Begin VB.TextBox txtRedCrd 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         TabIndex        =   6
         Top             =   4005
         Width           =   2370
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   5760
         ScaleHeight     =   1455
         ScaleWidth      =   1215
         TabIndex        =   5
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
      Begin Project1.desButton cmdBrowse 
         Height          =   735
         Left            =   5880
         TabIndex        =   18
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
         cBack           =   -2147483633
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
         TabIndex        =   30
         Top             =   3675
         Width           =   1800
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
         TabIndex        =   29
         Top             =   2280
         Width           =   1335
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   2955
         Width           =   1455
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
         TabIndex        =   26
         Top             =   855
         Width           =   1215
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
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
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
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
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
         TabIndex        =   23
         Top             =   2280
         Width           =   1215
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
         TabIndex        =   22
         Top             =   3000
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
         TabIndex        =   21
         Top             =   4440
         Width           =   1455
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
         TabIndex        =   20
         Top             =   3720
         Width           =   1215
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
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   7215
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   5640
         TabIndex        =   1
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
      Begin Project1.desButton cmdReset 
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Reset"
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
      Begin Project1.desButton cmdSave 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6840
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Player 
      Height          =   390
      Left            =   120
      Top             =   6120
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
      Left            =   2280
      Top             =   6120
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
End
Attribute VB_Name = "AddPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim picName As String
Dim spicname As String
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

    Unload AddPlayer
    
End Sub

Private Sub cmdReset_Click()

      txtName.Text = ""
      txtPId.Text = ""
      txtPReg.Text = ""
      txtPPos.Text = ""
      txtPDob.Text = ""
      txtPYelCrd.Text = ""
      txtPState.Text = ""
      txtPDoj.Text = ""
      txtRedCrd.Text = ""
      txtPtfrom.Text = ""
      txtPStatus.Text = ""
      txtPClub.Text = ""

End Sub

Private Sub cmdSave_Click()
        
    'On Error GoTo ErrHandler 'handles expected unsuccessful data entry error

       'Player.Refresh
       'Player.Recordset.Find ("ID_No = '" & Trim(txtPId.Text) & "'")

        If Trim(txtName.Text) = "" Or Trim(txtPId.Text) = "" Or Trim(txtPReg.Text) = "" _
            Or Trim(txtPPos.Text) = "" Or Trim(txtPDob.Text) = "" Or Trim(txtPYelCrd.Text) = "" _
            Or Trim(txtPState.Text) = "" Or Trim(txtPDoj.Text) = "" Or Trim(txtRedCrd.Text) = "" _
            Or Trim(txtPtfrom.Text) = "" Or Trim(txtPStatus.Text) = "" Or Trim(txtPClub.Text) = "" Then
            
            MsgBox "Required field missing. Please fill up ALL the fields.", vbOKOnly + vbExclamation, "SysMan"
            
            'checks the missing field and focuses on it
            If Trim(txtName.Text) = "" Then
                txtName.Text = ""
                txtName.SetFocus
            ElseIf Trim(txtPId.Text) = "" Then
                txtPId.Text = ""
                txtPId.SetFocus
            ElseIf Trim(txtPReg.Text) = "" Then
                txtPReg.Text = ""
                txtPReg.SetFocus
            ElseIf Trim(txtPClub.Text) = "" Then
                txtPClub.Text = ""
                txtPClub.SetFocus
            ElseIf Trim(txtPPos.Text) = "" Then
                txtPPos.Text = ""
                txtPPos.SetFocus
            ElseIf Trim(txtPDob.Text) = "" Then
                txtPDob.Text = ""
                txtPDob.SetFocus
            ElseIf Trim(txtPState.Text) = "" Then
                txtPState.Text = ""
                txtPState.SetFocus
             ElseIf Trim(txtPDoj.Text) = "" Then
                txtPDoj.Text = ""
                txtPDoj.SetFocus
             ElseIf Trim(txtPtfrom.Text) = "" Then
                txtPtfrom.Text = ""
                txtPtfrom.SetFocus
             ElseIf Trim(txtPStatus.Text) = "" Then
                txtPStatus.Text = ""
                txtPStatus.SetFocus
            ElseIf Trim(txtPYelCrd.Text) = "" Then
                txtPYelCrd.Text = ""
                txtPYelCrd.SetFocus
            ElseIf Trim(txtRedCrd.Text) = "" Then
                txtRedCrd.Text = ""
                txtRedCrd.SetFocus
            End If
        
            Exit Sub
        End If
        
        'if all fields are ok then transfer data to the database
        'checks if the password typed is similar with the password confirmation
        'If Trim(txtPReg.Text) = Player.Recordset.Fields(2) Then
            
            'MsgBox "The Player Registration No. is Already Exist." & vbCrLf + vbCrLf & "Please Check The ID No.", "System Admin"
            'txtPReg.Text = ""
            'txtPReg.SetFocus
            'Exit Sub
        'Else
        
        If IcValid(Trim(txtPId.Text)) = True Then
            MsgBox "The Player Id Is Already Exist " & vbCrLf & "Please Provide A Valid Id", vbInformation, "SysMan"
            txtPId.SetFocus
            SendKeys highLig
            Exit Sub
        End If
        
        adoplayer.RecordSource = "SELECT * FROM Players WHERE Club = '" & Trim(txtPClub.Text) & "'"
        adoplayer.Refresh

        If adoplayer.Recordset.RecordCount > 24 Then
            MsgBox "The Team Only Can Occupied Maximum 25 Members" & vbCrLf & "You Have Reached The Maximum Entry", vbInformation, "FAS System"
            Exit Sub
        End If
        
            Player.Refresh
            Player.Recordset.AddNew
        
            With Player.Recordset
                .Fields(0) = txtName.Text
                .Fields(1) = txtPId.Text
                .Fields(2) = txtPReg.Text
                .Fields(3) = txtPClub.Text
                .Fields(4) = txtPPos.Text
                .Fields(5) = txtPDob.Text
                .Fields(6) = txtPState.Text
                .Fields(7) = txtPDoj.Text
                .Fields(8) = txtPtfrom.Text
                .Fields(9) = txtPStatus.Text
                .Fields(10) = txtPYelCrd.Text
                .Fields(11) = txtRedCrd.Text
            End With
                        
            On Error Resume Next
            If mpicname <> "" Then
                Player.Recordset.Fields(12) = mpicname
            End If
            
            Player.Recordset.Update
            
            'confirms that data has already been entered to the database
            Player.Refresh
                        
            Player.Recordset.MoveFirst 'will generate an error if data has not been entered
            
            MsgBox "Entry Data Successfull !!!", vbInformation, "SysMan"
            
            Unload AddPlayer

            Exit Sub
                        
        'End If
        
'Exit Sub

'ErrHandler:
    'MsgBox "Player Registration Number already exists. Please Choose An Appropriate Registration Number", vbOKOnly, "System Admin"
    'txtPReg.SetFocus
    'SendKeys highLig
    'Exit Sub
        
End Sub

Private Sub Form_Load()

    Call DataConn(Player, "Players")
    Call DataConn(Club, "Clubs")
    
    AddPlayer.txtPClub.Text = frmEditClub.txtsearchname.Text
    
End Sub
