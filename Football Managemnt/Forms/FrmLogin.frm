VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  Welcome to Football Management System"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoLig 
      Height          =   375
      Left            =   1680
      Top             =   6240
      Width           =   1200
      _ExtentX        =   2117
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
   Begin Project1.desButton cmdClerk 
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "Clerk Login"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin Project1.desButton cmdAdmin 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "Admin Login"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin MSAdodcLib.Adodc ADOLog 
      Height          =   495
      Left            =   120
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Football Management System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   -120
      Width           =   7425
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   6090
      Left            =   -30
      Picture         =   "FrmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   7605
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadmin_Click()
        
        FrmLogin.Hide
        Unload FrmLogin
    AdminLogin.Show
       
End Sub

Private Sub cmdClerk_Click()
        
        FrmLogin.Hide
        Unload FrmLogin
    ClerkLogin.Show
         
End Sub

Private Sub Form_Load()

   On Error GoTo ErrHandler 'we will use an intentional error to facilitate new user input

    frmSplash.lblIni.Caption = "Accessing database..."
       
       Call DataConn(ADOLog, "Admin")
       Call DataConn(AdoLig, "Clerk")

    frmSplash.lblIni.Caption = "Initialization complete!"
    
        'refreshes database status
    ADOLog.Refresh
    AdoLig.Refresh
    
    'intentionally create an error situation if there is no record in the Dbase
    ADOLog.Recordset.MoveFirst
    AdoLig.Recordset.MoveFirst
    
    'FrmLogin.Show

Exit Sub


ErrHandler: 'handles the error by asking the new user to input initial settings
    'frmSplash.Hide
    
    'Load frmAdminSetup
    
    'frmSplash.lblIni.Caption = "Preparing initial setup..."
    
    'frmAdminSetup.Show vbModal
        
    'FrmLogin.Show
    
    ADOLog.Refresh
    AdoLig.Refresh


End Sub

