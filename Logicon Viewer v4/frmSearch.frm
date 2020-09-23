VERSION 5.00
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Search Parameters"
   ClientHeight    =   4245
   ClientLeft      =   1650
   ClientTop       =   3645
   ClientWidth     =   4785
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4245
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Replace String (All Matches)"
      Height          =   345
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2790
      Width           =   2835
   End
   Begin VB.TextBox txtReplace 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   180
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start From Current Position"
      Default         =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3690
      Width           =   2835
   End
   Begin VB.OptionButton optSearchMode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Long (QByte)"
      Height          =   375
      Index           =   4
      Left            =   2340
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2265
   End
   Begin VB.OptionButton optSearchMode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Short (DByte)"
      Height          =   525
      Index           =   3
      Left            =   2340
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   510
      Width           =   2265
   End
   Begin VB.OptionButton optSearchMode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Byte"
      Height          =   375
      Index           =   2
      Left            =   2340
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   2265
   End
   Begin VB.OptionButton optSearchMode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text (Double Byte)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   570
      Width           =   2115
   End
   Begin VB.OptionButton optSearchMode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text (Single Byte)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   180
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3690
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start From Beginning"
      Height          =   345
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2835
   End
   Begin VB.TextBox txtSearchString 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Top             =   1830
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(use &&H for HEX values)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   2340
      TabIndex        =   11
      Top             =   1470
      Width           =   2280
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OK_Focus As Boolean

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSearch_Click(Index As Integer)
  If Index = 2 Then
    If Len(txtSearchString.Text) <> Len(txtReplace.Text) Or Len(txtSearchString) = 0 Then
      MsgBox "When replacing text strings the strings MUST be of equal length and can not be empty!", vbOKOnly, "ERROR"
      
      If OK_Focus Then txtSearchString.SetFocus
      
      Exit Sub
    Else
      If MsgBox("You are about to replace (" & txtSearchString.Text & ") with (" & txtReplace.Text & ")" & vbNewLine & vbNewLine & "This action can NOT be undone. Are you sure you wish to proceed?", vbYesNo, "Proceed?") <> vbYes Then
        If OK_Focus Then txtSearchString.SetFocus
        
        Exit Sub
      End If
    End If
  End If
  
  SearchString = txtSearchString.Text
  ReplaceString = txtReplace.Text
  
  SearchStart = Index + 1
  
  Unload Me
End Sub

Private Sub Form_Load()
  SizeForm Me
  
  OK_Focus = False
  
  Me.Left = (Screen.Width - Me.Width) \ 2
  Me.Top = (Screen.Height - Me.Height) \ 2
  
  optSearchMode(SearchMode).Value = 1
  txtSearchString.Text = SearchString
  txtSearchString.SelLength = Len(SearchString)
  txtReplace.Text = ReplaceString
  txtReplace.SelLength = Len(SearchString)
  
  If SearchMode < 2 Then
    txtReplace.Visible = True
    cmdSearch(2).Visible = True
  Else
    txtReplace.Visible = False
    cmdSearch(2).Visible = False
  End If
  
  SearchStart = 0
  OK_Focus = True
End Sub

Private Sub optSearchMode_Click(Index As Integer)
  SearchMode = Index
  
  If SearchMode < 2 Then
    txtReplace.Visible = True
    cmdSearch(2).Visible = True
  Else
    txtReplace.Visible = False
    cmdSearch(2).Visible = False
  End If
  
  If OK_Focus Then txtSearchString.SetFocus
End Sub
