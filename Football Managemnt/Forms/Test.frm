VERSION 5.00
Begin VB.Form Test 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "Report"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_Click()

    Dim ad_Name As String
    Dim sSql As String
    
        ad_Name = InputBox("Enter The Club Name :", "SysMan")
        
        If Len(ad_Name) > 0 Then
                            
            DataEnvironment1.rsstartlist.Open
            DataEnvironment1.rsstartlist.Filter = ""
            DataEnvironment1.rsstartlist.Filter = "Club = '" & ad_Name & "'"
            CheckList.Show vbModal, Me
            CheckList.Sections(3).Visible = False
            DataEnvironment1.rsstartlist.Close
        Else
            DataEnvironment1.rsstartlist.Open
            DataEnvironment1.rsstartlist.Filter = ""
            CheckList.Show vbModal, Me
            CheckList.Sections(3).Visible = False
            DataEnvironment1.rsstartlist.Close

        End If
            
            
End Sub


