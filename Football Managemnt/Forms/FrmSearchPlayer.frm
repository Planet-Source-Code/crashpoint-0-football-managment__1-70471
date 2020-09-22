VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmSearchPlayer 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  Player Search"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   5640
      TabIndex        =   30
      Top             =   240
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Search 
      Height          =   390
      Left            =   960
      Top             =   6960
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
   Begin VB.TextBox txtsearchname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   1
      Top             =   180
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
      TabIndex        =   2
      Top             =   1440
      Width           =   7200
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   5640
         ScaleHeight     =   1455
         ScaleWidth      =   1215
         TabIndex        =   28
         Top             =   1320
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
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4005
         Width           =   2370
      End
      Begin VB.TextBox txtPYelCrd 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4725
         Width           =   2370
      End
      Begin VB.TextBox txtPStatus 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3285
         Width           =   2370
      End
      Begin VB.TextBox txtPState 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2520
         Width           =   2370
      End
      Begin VB.TextBox txtPReg 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1845
         Width           =   2370
      End
      Begin VB.TextBox txtPId 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1845
         Width           =   2370
      End
      Begin VB.TextBox txtPClub 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1155
         Width           =   2370
      End
      Begin VB.TextBox txtPDoj 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3255
         Width           =   2370
      End
      Begin VB.TextBox txtPPos 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1155
         Width           =   2370
      End
      Begin VB.TextBox txtPDob 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   2370
      End
      Begin VB.TextBox txtPtfrom 
         BackColor       =   &H80000009&
         DataSource      =   "Search"
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
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3975
         Width           =   2370
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Image Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   29
         Top             =   2880
         Width           =   975
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
         TabIndex        =   23
         Top             =   4440
         Width           =   1575
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   12
         Top             =   855
         Width           =   1335
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
         TabIndex        =   11
         Top             =   2955
         Width           =   1455
      End
      Begin VB.Label lblname 
         BackColor       =   &H80000009&
         Caption         =   "Player name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Width           =   2700
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
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
         TabIndex        =   6
         Top             =   3675
         Width           =   1920
      End
   End
   Begin MSAdodcLib.Adodc Club 
      Height          =   390
      Left            =   3240
      Top             =   6960
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
      TabIndex        =   26
      Top             =   840
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "New Search"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin Project1.desButton cmdBack 
      Height          =   495
      Left            =   2640
      TabIndex        =   27
      Top             =   840
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "Close"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
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
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "FrmSearchPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim team As String

Private Sub cmdBack_Click()

    Unload FrmSearchPlayer
    
End Sub

Private Sub cmdNewSearch_Click()
    
        txtsearchname.Text = ""
        lblName.Caption = ""
    
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
        
End Sub

Private Sub cmdSearch_Click()

Dim pic As String
Dim CPic As String
Dim temp As String
    
On Error GoTo NotFound

    
    If Trim(txtsearchname.Text) = "" Then
        MsgBox "Please Enter Appropriate Value", vbCritical, "System Admin"
        Exit Sub
    End If
            
        'temp = Search.Recordset.Fields(1)

        Search.Refresh
        Search.Recordset.Find ("Club = '" & Trim(txtsearchname.Text) & "'")
        
        Club.Refresh
        Club.Recordset.Find ("Club_Name ='" & Trim(txtsearchname.Text) & "'")

        

        On Error Resume Next
        
        lblName.Caption = Search.Recordset.Fields("Name")
        txtPId.Text = Search.Recordset.Fields("ID_No")
        txtPReg.Text = Search.Recordset.Fields("Reg_No")
        
        txtPClub.Text = Search.Recordset.Fields("Club")
        txtPPos.Text = Search.Recordset.Fields("Position")
        txtPDob.Text = Search.Recordset.Fields("DOB")
        txtPState.Text = Search.Recordset.Fields("State")
        txtPDoj.Text = Search.Recordset.Fields("DOJ")
        txtPtfrom.Text = Search.Recordset.Fields("TFrom")
        txtPStatus.Text = Search.Recordset.Fields("Status")
        txtPYelCrd.Text = Search.Recordset.Fields("Yellow_Crd")
        txtRedCrd.Text = Search.Recordset.Fields("Red_Crd")
        
        pic = Search.Recordset.Fields("Picture")
        'CPic = Club.Recordset.Fields("Logo")
            imageplayer.Picture = LoadPicture(pic)
        
        txtsearchname.SetFocus
        SendKeys HiLyt
        
        Exit Sub
NotFound:
        MsgBox "The player profile you requested could not be found.", vbOKOnly + vbExclamation, "System Admin"
        
        
        txtsearchname.SetFocus
        SendKeys HiLyt

End Sub

Private Sub Form_Load()

 On Error GoTo ErrHandle
    
    Call DataConn(Search, "Players")
    Call DataConn(Club, "Clubs")
    
    
  Exit Sub
        
ErrHandle:


End Sub

