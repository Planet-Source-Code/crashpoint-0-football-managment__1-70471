VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EditPlayer 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  Edit Player Information"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6240
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
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
      TabIndex        =   6
      Top             =   1440
      Width           =   7200
      Begin VB.TextBox txtName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   5055
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
         TabIndex        =   18
         Top             =   3975
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
         TabIndex        =   17
         Top             =   2520
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   3255
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
         TabIndex        =   14
         Top             =   1155
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
         TabIndex        =   13
         Top             =   1845
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
         TabIndex        =   12
         Top             =   1845
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
         TabIndex        =   11
         Top             =   2520
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
         TabIndex        =   9
         Top             =   4725
         Width           =   2370
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   32
         Top             =   2760
         Width           =   975
         _extentx        =   1720
         _extenty        =   1296
         caption         =   "Browse"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         enabled         =   0   'False
         cback           =   -2147483633
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   3720
         Width           =   1215
      End
   End
   Begin VB.TextBox txtsearchname 
      Height          =   390
      Left            =   1920
      TabIndex        =   5
      Top             =   180
      Width           =   3360
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   7215
      Begin Project1.desButton cmdClose 
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
         caption         =   "Close"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin Project1.desButton cmdUpdate 
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
         caption         =   "Update"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         enabled         =   0   'False
         cback           =   -2147483633
      End
      Begin Project1.desButton cmdReload 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "Reload"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
   End
   Begin Project1.desButton cmdSearch 
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "Search"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin MSAdodcLib.Adodc Search 
      Height          =   390
      Left            =   1440
      Top             =   7680
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
      Top             =   7680
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
      TabIndex        =   30
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
      TabIndex        =   31
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "EditPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

End Sub

Private Sub cmdBrowse_Click()

    Dim sPic As String
    Dim ssPic As String
    
      dlgCommon.Filter = "Picture Files(*.jpg; *.bmp; *.gif)|*.jpg;*.bmp;*.gif"
      dlgCommon.ShowOpen
      
        sPic = dlgCommon.FileName
        ssPic = Mid$(sPic, InStrRev(sPic, "/") + 1)
        
            If sPic <> "" Then
                imageplayer.Picture = LoadPicture(sPic)
            End If

End Sub

Private Sub cmdClose_Click()

    Unload EditPlayer
    
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
        'imageclub.Picture = LoadPicture("")
        
End Sub

Private Sub cmdReload_Click()

    If MsgBox("This will reload the current admin profile !!!." & vbCrLf & _
            " Any unsaved data will be lost. Do You Want To Proceed?", _
                            vbYesNo + vbQuestion, "SysMan") = vbYes Then
        
        imageplayer.Picture = LoadPicture("")
        dlgCommon.FileName = ""
        
        Search.Refresh
        Search.Recordset.Find ("Reg_No = '" & Trim(txtsearchname.Text) & "'")
         
        On Error Resume Next
        
        txtName.Text = Search.Recordset.Fields("Name")
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
        
        On Error Resume Next
            imageplayer.Picture = LoadPicture("")
            
        txtName.Enabled = True
        txtPId.Enabled = True
        txtPReg.Enabled = True
        txtPClub.Enabled = True
        txtPPos.Enabled = True
        txtPDob.Enabled = True
        txtPState.Enabled = True
        txtPDoj.Enabled = True
        txtPtfrom.Enabled = True
        txtPStatus.Enabled = True
        txtPYelCrd.Enabled = True
        txtRedCrd.Enabled = True
        
        txtPId.Locked = False
        txtPReg.Locked = False
        txtPClub.Locked = False
        txtPPos.Locked = False
        txtPDob.Locked = False
        txtPState.Locked = False
        txtPDoj.Locked = False
        txtPtfrom.Locked = False
        txtPStatus.Locked = False
        txtPYelCrd.Locked = False
        txtRedCrd.Locked = False
                 
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
    
        Search.Refresh
        Search.Recordset.Find ("ID_No = '" & Trim(txtsearchname.Text) & "'")
         
        On Error Resume Next
        
        txtName.Text = Search.Recordset.Fields("Name")
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
            
            On Error Resume Next
            imageplayer.Picture = LoadPicture(pic)
        
        txtsearchname.SetFocus
        SendKeys HiLyt
        
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
        
        On Error Resume Next

        Search.Recordset.Fields("Name") = txtName.Text
        Search.Recordset.Fields("ID_No") = txtPId.Text
        Search.Recordset.Fields("Reg_No") = txtPReg.Text
        Search.Recordset.Fields("Club") = txtPClub.Text
        Search.Recordset.Fields("Position") = txtPPos.Text
        Search.Recordset.Fields("DOB") = txtPDob.Text
        Search.Recordset.Fields("State") = txtPState.Text
        Search.Recordset.Fields("DOJ") = txtPDoj.Text
        Search.Recordset.Fields("TFrom") = txtPtfrom.Text
        Search.Recordset.Fields("Status") = txtPStatus.Text
        Search.Recordset.Fields("Yellow_Crd") = txtPYelCrd.Text
        Search.Recordset.Fields("Red_Crd") = txtRedCrd.Text
             
        On Error Resume Next
        If dlgCommon.FileName <> "" Then
            Search.Recordset.Fields("Picture") = dlgCommon.FileName
        End If
        
        Search.Recordset.Update
        Search.Refresh
            
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

    Call DataConn(Search, "Players")
    Call DataConn(Club, "Clubs")

End Sub


