VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   1110
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmMenu.frx":0000
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4500
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
      Begin VB.Menu MnuNewBor 
         Caption         =   "New Admin Profile"
      End
      Begin VB.Menu mnuNewLib 
         Caption         =   "New Club Profile"
      End
      Begin VB.Menu mnuClerk 
         Caption         =   "New Clerk Profile"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditBor 
         Caption         =   "Edit Admin Profile"
      End
      Begin VB.Menu mnuEditLib 
         Caption         =   "Edit Club Profile"
      End
      Begin VB.Menu mnuEditClerk 
         Caption         =   "Edit Clerk Profile"
      End
      Begin VB.Menu mnuEditPlayer 
         Caption         =   "Edit Player Profile"
      End
   End
   Begin VB.Menu mnuDel 
      Caption         =   "Delete"
      Begin VB.Menu mnuDelBor 
         Caption         =   "Delete Admin Profile"
      End
      Begin VB.Menu mnuDelLib 
         Caption         =   "Delete Club Profile"
      End
      Begin VB.Menu mnuDelClerk 
         Caption         =   "Delete Clerk Profile"
      End
      Begin VB.Menu mnuDelPlayer 
         Caption         =   "Delete Player Profile"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Begin VB.Menu mnuSeaPlayer 
         Caption         =   "Search Player"
      End
      Begin VB.Menu mnuSeaClub 
         Caption         =   "Search Club"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuClerk_Click()

    On Error Resume Next
        ClerkNew.Show vbModal
    
End Sub

Private Sub mnuDelBor_Click()
On Error Resume Next
    
    Confirm.Show vbModal
    'AdminDel.Show vbModal
    
    'MsgBox "In Progress"
End Sub

Private Sub mnuDelClerk_Click()
On Error Resume Next
    Confirm3.Show vbModal
    
End Sub

Private Sub mnuDelLib_Click()
On Error Resume Next
    
    Confirm2.Show vbModal
    'HNDStudDel.Show vbModal
    
    'MsgBox "In Progress"
End Sub

Private Sub mnuDelPlayer_Click()

    On Error Resume Next
        Confirm4.Show vbModal
        
End Sub

Private Sub mnuEditBor_Click()
On Error Resume Next
    
    AdminEdit.Show vbModal
End Sub

Private Sub mnuEditClerk_Click()

    On Error Resume Next
        ClerkEdit.Show vbModal
        
End Sub

Private Sub mnuEditLib_Click()
On Error Resume Next
    frmEditClub.Show vbModal
End Sub

Private Sub mnuEditPlayer_Click()

    On Error Resume Next
        EditPlayer.Show vbModal
        
End Sub

Private Sub MnuNewBor_Click()
On Error Resume Next
    
    AdminNew.Show vbModal
End Sub

Private Sub mnuNewLib_Click()
On Error Resume Next
    
    frmNewClub.Show vbModal
End Sub

Private Sub mnuSeaClub_Click()

    On Error Resume Next
        
        FrmSearchCountry.Show vbModal
        
End Sub

Private Sub mnuSeaPlayer_Click()

    On Error Resume Next
        
        FrmSearchPlayer.Show vbModal
        
End Sub
