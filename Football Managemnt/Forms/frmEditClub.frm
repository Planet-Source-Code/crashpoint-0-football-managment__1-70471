VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditClub 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :.  Edit Club Information"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2040
      TabIndex        =   8
      Top             =   165
      Width           =   3360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Player Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   8220
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin Project1.desButton cmdAdd 
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "Add New Player"
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
      Height          =   1680
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   8175
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6480
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.Image imgClub 
         Height          =   1455
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1455
      End
   End
   Begin Project1.desButton cmdBack 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
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
   Begin Project1.desButton cmdnewsearch 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
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
   Begin Project1.desButton cmdSearch 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc adoFilter2 
      Height          =   375
      Left            =   120
      Top             =   7560
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
   Begin MSAdodcLib.Adodc ClubInfo 
      Height          =   375
      Left            =   3960
      Top             =   7560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc adoFilter 
      Height          =   375
      Left            =   1800
      Top             =   7560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc PlayerInfo 
      Height          =   375
      Left            =   6240
      Top             =   7560
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
      Caption         =   "PlayerInfo"
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
   Begin Project1.desButton cmdedit 
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Edit"
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
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Save"
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
   Begin Project1.desButton cmdcan 
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Cancel"
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
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frmEditClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim delete As Boolean
Dim Confirm As Integer

Private Sub cmdAdd_Click()

    AddPlayer.Show vbModal
    
        AddPlayer.txtPClub.Text = frmEditClub.txtsearchname.Text
        
End Sub

Private Sub cmdBack_Click()

    Unload frmEditClub
    
End Sub

Private Sub cmdcan_Click()

    PlayerInfo.Refresh
    ClubInfo.Refresh
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
    cmdSave.Enabled = False
    cmdedit.Enabled = True

End Sub

Private Sub cmdedit_Click()

    cmdcan.Enabled = True
    cmdSave.Enabled = True
    cmdedit.Enabled = False
    
    DataGrid1.AllowUpdate = True
    DataGrid1.Enabled = True
    DataGrid2.AllowUpdate = True
    DataGrid2.Enabled = True
        
End Sub

Private Sub cmdNewSearch_Click()

    txtsearchname.Text = ""
    txtsearchname.SetFocus
    DataGrid1.ClearFields
    DataGrid2.ClearFields
    imgClub.Picture = LoadPicture("")
    FrmSearchCountry.Refresh
    
End Sub

Private Sub cmdSave_Click()

    If delete = True Then
        Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
        If Confirm = vbYes Then
            ClubInfo.Recordset.delete
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
    cmdSave.Enabled = False
    cmdedit.Enabled = True
    
End Sub

Private Sub cmdSearch_Click()
                       
On Error GoTo NotFound

    Dim temp As String
    Dim temp2 As String
    Dim pic As String
    Dim counter As Integer
               
        If Trim(txtsearchname.Text) = "" Then
            MsgBox "Please Enter The Appropriate Data", vbInformation, "Sys Man"
            txtsearchname.SetFocus
            SendKeys highLig
            Exit Sub
        End If

        PlayerInfo.Refresh
        PlayerInfo.Recordset.Find ("Club = '" & Trim(txtsearchname.Text) & "'")
        
        ClubInfo.Refresh
        ClubInfo.Recordset.Find ("Club_Name ='" & Trim(txtsearchname.Text) & "'")
        
        DataGrid1.ClearFields
        DataGrid2.ClearFields
        txtsearchname.SetFocus
            
            temp = PlayerInfo.Recordset.Fields(3)
            temp2 = ClubInfo.Recordset.Fields(1)

            On Error Resume Next
            Call BorGrid
                pic = ClubInfo.Recordset.Fields("Logo")
                imgClub.Picture = LoadPicture(pic)
                
        SendKeys highLig
    Exit Sub

NotFound:
    MsgBox "The record you requested could not be found.", vbOKOnly + vbExclamation, "System Admin"
    txtsearchname.SetFocus
    SendKeys highLig

End Sub

Public Sub BorGrid()

On Error Resume Next
 'filter the books to show only the ones borrowed by the Current Borrower
    adoFilter.RecordSource = "SELECT * FROM Players WHERE Club = '" & Trim(txtsearchname.Text) & "'"
    adoFilter.Refresh
    
    adoFilter2.RecordSource = "SELECT * FROM Clubs WHERE Club_Name = '" & Trim(txtsearchname.Text) & "'"
    adoFilter2.Refresh
    
    Set DataGrid1.DataSource = adoFilter
    Set DataGrid2.DataSource = adoFilter2


    With DataGrid2
        .Columns(0).DataField = "Manager"
        .Columns(0).Caption = "Name"
        .Columns(0).Width = 1500
        
        .Columns(1).DataField = "Club_Name"
        .Columns(1).Caption = "Club Name"
        .Columns(1).Width = 2000

        .Columns(2).DataField = "Founded"
        .Columns(2).Caption = "Year Founded"
        .Columns(2).Width = 2000
        
        .Columns(4).DataField = "League"
        .Columns(4).Caption = "League"
        .Columns(4).Width = 1300
        
    End With
    
    With DataGrid1
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
                
        .Columns(7).DataField = "DOJ"
        .Columns(7).Caption = "Date of Join"
        .Columns(7).Width = 1500
        
        .Columns(8).DataField = "TFrom"
        .Columns(8).Caption = "Transffered From"
        .Columns(8).Width = 1500
        
        .Columns(9).DataField = "Status"
        .Columns(9).Caption = "Status"
        .Columns(9).Width = 1500
        
        .Columns(10).DataField = "Yellow_Crd"
        .Columns(10).Caption = "Yellow Card"
        .Columns(10).Width = 0
        
        .Columns(11).DataField = "Red_Crd"
        .Columns(11).Caption = "Red Card"
        .Columns(11).Width = 0
        
        .Columns(12).DataField = "Picture"
        .Columns(12).Caption = "Picture"
        .Columns(12).Width = 0
        
     End With
     
      'imgClub.Picture = LoadPicture("clubinfo.Recordset.Fields(9)")
                                    
End Sub

Private Sub Form_Load()

  On Error GoTo ErrHandle
  
    Call DataConn(PlayerInfo, "Players")
    Call DataConn(ClubInfo, "Clubs")
        
    adoFilter.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
        
     'Sets the command type to Table
        adoFilter.CommandType = adCmdText
    
    adoFilter2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\football.mdb;Persist Security Info=False; Jet OLEDB:Database Password = crimson119"
        
     'Sets the command type to Table
        adoFilter2.CommandType = adCmdText

    DataGrid2.ClearFields
    
    PlayerInfo.Refresh
    ClubInfo.Refresh
    PlayerInfo.Recordset.MoveFirst
    ClubInfo.Recordset.MoveFirst

  Exit Sub
        
ErrHandle:

End Sub


