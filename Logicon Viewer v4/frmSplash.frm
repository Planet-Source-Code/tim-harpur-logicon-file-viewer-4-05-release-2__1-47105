VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4260
      Top             =   5250
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2070
      TabIndex        =   1
      Top             =   6750
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   202
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   314
      TabIndex        =   0
      Top             =   210
      Width           =   4710
   End
   Begin VB.Label lblProgram 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   5490
      Width           =   4605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSplash.frx":309E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   240
      TabIndex        =   2
      Top             =   4140
      Width           =   4590
   End
   Begin VB.Image imgTop 
      Height          =   24300
      Left            =   -2190
      Picture         =   "frmSplash.frx":3192
      Top             =   6570
      Visible         =   0   'False
      Width           =   4590
   End
   Begin VB.Image imgBottom 
      Height          =   24900
      Left            =   2700
      Picture         =   "frmSplash.frx":1B93E
      Top             =   6600
      Visible         =   0   'False
      Width           =   4590
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private ProgActive As Long

Public TitleBar As String

Private Sub cmdDone_Click()
  If ProgActive = 1 Then ProgActive = 0
End Sub

Private Sub Form_Load()
  lblProgram.Caption = TitleBar
  ProgActive = 1
  
  Timer1.Enabled = True
End Sub

Private Sub RenderFrame()
  Static CurrentFrame As Long
  
  CurrentFrame = CurrentFrame + 1
  If CurrentFrame >= 20 Then CurrentFrame = 0
  
  Picture1.PaintPicture imgTop.Picture, 4, 4, 306, 81, 0, CurrentFrame * 81, 306, 81
  Picture1.PaintPicture imgBottom.Picture, 4, 169, 306, 83, 0, CurrentFrame * 83, 306, 83
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ProgActive <> -1 Then
    Cancel = 1
    
    ProgActive = 0
  End If
End Sub

Private Sub Timer1_Timer()
  Dim lastTime As Long
  
  Timer1.Enabled = False
  
  Do While ProgActive > 0
    lastTime = WaitTime(lastTime, 10) + 80
    
    RenderFrame
  Loop
  
  ProgActive = -1
  
  Unload Me
End Sub

Private Function WaitTime(returnTime As Long, Optional MaxCarryOver As Long = 0, Optional ByVal UseRelativeTime As Boolean = False)
  Dim CarryOver As Long
  
  WaitTime = timeGetTime()
  If UseRelativeTime Then returnTime = WaitTime + returnTime
  
  Do While WaitTime < returnTime
    DoEvents
    
    WaitTime = timeGetTime()
  Loop
  
  CarryOver = WaitTime - returnTime
  If CarryOver > MaxCarryOver Then CarryOver = MaxCarryOver
  
  If CarryOver > 0 Then WaitTime = WaitTime - CarryOver
End Function
