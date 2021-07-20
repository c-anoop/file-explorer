VERSION 5.00
Begin VB.Form frmExplorer 
   Caption         =   "Explorer"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmExplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filExplorer 
      Height          =   2235
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.DirListBox dirExplorer 
      Height          =   2115
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.DriveListBox drvExplorer 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblFolders 
      AutoSize        =   -1  'True
      Caption         =   "Folders"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   570
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dirExplorer_Change()

 ' This subroutine sets the path for files
   filExplorer.Path = dirExplorer.Path
     
End Sub

Private Sub drvExplorer_Change()
 
 ' This subroutine sets the path for Folders
   On Error GoTo errhandler
   dirExplorer.Path = drvExplorer.Drive
   Exit Sub
errhandler: _
               Call MsgBox("Error Reading Drive", _
                vbCritical + vbOKOnly, "Error")
    drvExplorer.Refresh
End Sub



Private Sub Form_Resize()
 
 ' This Subroutine sets the window size
   On Error Resume Next
   
   drvExplorer.Width = frmExplorer.Width - 1200
   dirExplorer.Width = frmExplorer.Width / 3 + 700
   If frmExplorer.Width > dirExplorer.Width + 950 Then
      filExplorer.Width = frmExplorer.Width - _
      dirExplorer.Width - 950
   End If
   filExplorer.Left = dirExplorer.Width + 600
   If filExplorer.Height > 100 Or frmExplorer.Height > 3000 Then
      dirExplorer.Height = frmExplorer.Height - 2500
      filExplorer.Height = frmExplorer.Height - 2500
   End If

End Sub

