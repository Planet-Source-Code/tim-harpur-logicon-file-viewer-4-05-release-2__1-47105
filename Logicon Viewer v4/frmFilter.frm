VERSION 5.00
Begin VB.Form frmFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Set File Filter"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2400
      TabIndex        =   2
      Top             =   750
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   855
   End
   Begin VB.ComboBox cmbList 
      Height          =   315
      ItemData        =   "frmFilter.frx":0000
      Left            =   150
      List            =   "frmFilter.frx":0034
      TabIndex        =   0
      Top             =   210
      Width           =   3105
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  frmMain.Explorer1.FileFilter = cmbList.Text
  
  Unload Me
End Sub

Private Sub Form_Load()
  SizeForm Me
  
  If frmMain.Explorer1.FileFilter = "" Then
    cmbList.Text = "*.*"
  Else
    cmbList.Text = frmMain.Explorer1.FileFilter
  End If
End Sub
