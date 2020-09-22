VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ClerkLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :. Clerk Login"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      Begin VB.TextBox txtAdminId 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtAdminPass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin Project1.desButton cmdExit 
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         caption         =   "Exit"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin Project1.desButton cmdOk 
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
         caption         =   "Ok"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin VB.Label lblLogAttem 
         BackColor       =   &H80000009&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Login Attempt :-"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Clerk Password :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Clerk Name       :"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc adoLogs 
      Height          =   330
      Left            =   2640
      Top             =   3240
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Admin Logs"
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
   Begin MSAdodcLib.Adodc AdoCollege 
      Height          =   330
      Left            =   1320
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Inst Login"
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
   Begin MSAdodcLib.Adodc AdoLogin 
      Height          =   330
      Left            =   0
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DBase Login"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "CLERK LOGIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "ClerkLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public log_attem As String

Private Sub cmdexit_Click()

    If MsgBox("Do You Want To Switch User", vbYesNo, "SysMan") = vbYes Then
        Unload ClerkLogin
        FrmLogin.Show
    Else
        Exit Sub
    End If
    
End Sub

Private Sub cmdOk_Click()
    
    On Error GoTo NotFound
    
    AdoLogin.Refresh
    
   
    AdoLogin.Recordset.Find ("Username = '" & txtAdminId.Text & "'")
    AdoLogin.Recordset.Find ("Password ='" & txtAdminPass.Text & "'")
    
    tempUser2 = AdoLogin.Recordset.Fields("Username")
    tempPass2 = AdoLogin.Recordset.Fields("Password")
            
    If tempUser2 = txtAdminId.Text Then
        If tempPass2 = txtAdminPass.Text Then
            
           On Error Resume Next
            CName = AdoLogin.Recordset.Fields("Name")
            AId = AdoLogin.Recordset.Fields("ID_No")
            AUsername = AdoLogin.Recordset.Fields("Username")
            CPass = AdoLogin.Recordset.Fields("Password")
            AAdd1 = AdoLogin.Recordset.Fields("Address1")
            AAdd2 = AdoLogin.Recordset.Fields("Address2")
            APoscode = AdoLogin.Recordset.Fields("Poscode")
            ACity = AdoLogin.Recordset.Fields("City")
            AState = AdoLogin.Recordset.Fields("State")
            ACountry = AdoLogin.Recordset.Fields("Country")
            ATel = AdoLogin.Recordset.Fields("Telphone")
            AEMail = AdoLogin.Recordset.Fields("E-Mail")
            APic = AdoLogin.Recordset.Fields("Picture")
            
            
            AUsername = tempUser
            APass = tempPass
                        
            frmInsignia.lblName.Caption = CName 'assigns the librarian name to FrmInsignia
            
            MDIMain.Show
            
            Unload ClerkLogin
            
            ClerkLogin.Hide
            Unload Me
            
            'Call Record_Logins(admin_id, Now)
            
           
        Else
        
            MsgBox "Invalid Password. Access Denied.", vbOKOnly + vbExclamation, "SysMan"
            txtAdminPass.SetFocus
            SendKeys highLig
            
        End If
    Else
    
            MsgBox "Invalid Username. Access Denied.", vbOKOnly + vbExclamation, "SysMan"
            txtAdminId.Text = ""
    End If
    
    Exit Sub
    
NotFound:
                                   
    log_attem = log_attem - 1
    lblLogAttem.Caption = log_attem
    
        If log_attem = 0 Then
            MsgBox "You Have Already Used all The Attempt Given", vbCritical, "SysMan"
            End
        Else
            MsgBox "User name or password does not exist. Please enter a valid data." & vbCrLf & "You Have Only " & log_attem & " Attempt Left", vbInformation, "SysMan"
            txtAdminId.SetFocus
            SendKeys highLig
            AdoLogin.Refresh
        End If
                                   

End Sub

Private Sub Form_Load()
   
On Error GoTo ErrHandler
    
        frmSplash.lblIni.Caption = "Accesing Database..."
        
        Call DataConn(AdoLogin, "Clerk") ' connect to admin db

        frmSplash.lblIni.Caption = "Intialization Complete..."
    
        AdoLogin.Refresh
            
        AdoLogin.Recordset.MoveFirst
        
        ClerkLogin.Show
        
        log_attem = 3
        lblLogAttem.Caption = log_attem
    
    Exit Sub
    
ErrHandler:
    
    'frmSplash.Hide
    Unload Me
    
    MsgBox "Its Seems That You Are Using This System For First Time.." & vbCrLf & _
        "Football Management System", vbInformation, "SysMan"
    
    'Unload AdminLogin
    ClerkLogin.Hide
    
    Load ClerkNew
    
    frmSplash.lblIni.Caption = "Preparing Initial Setup..."
    
    ClerkNew.Show vbModal
    
    ClerkLogin.Show
    log_attem = 3
    lblLogAttem.Caption = log_attem
    
    AdoLogin.Refresh


End Sub

Private Sub Form_Unload(Cancel As Integer)

    'AdoLogin.Recordset = Nothing
    txtAdminId.Text = ""
    txtAdminPass.Text = ""
    
End Sub

Private Sub txtAdminId_GotFocus()

    SendKeys highLig
    
End Sub

Private Sub txtAdminId_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtAdminPass.SetFocus
        SendKeys highLig
    End If
    
End Sub

Private Sub txtAdminPass_GotFocus()

    SendKeys highLig
    
End Sub

Private Sub txtAdminPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdOk_Click
    End If
    
End Sub


