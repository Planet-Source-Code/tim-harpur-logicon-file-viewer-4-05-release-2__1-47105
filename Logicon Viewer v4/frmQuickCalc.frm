VERSION 5.00
Begin VB.Form frmQuickCalc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   375
      Left            =   2813
      TabIndex        =   1
      Top             =   1500
      Width           =   1335
   End
   Begin VB.TextBox txtEquation 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   188
      TabIndex        =   0
      Top             =   900
      Width           =   6585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(supports +, -, and unsigned hexadecimal and binary #s - prefix with H or B)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   390
      TabIndex        =   3
      Top             =   510
      Width           =   6165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Enter Number or Basic Equation to Calculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1110
      TabIndex        =   2
      Top             =   150
      Width           =   4755
   End
End
Attribute VB_Name = "frmQuickCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub Command1_Click()
  Dim textEquation As String, textPos As Long, textLength As Long
  
  'calculate result
  
  textEquation = txtEquation.Text
  textLength = Len(textEquation)
  quickCalcResult = 0
  textPos = 1
  
  On Error GoTo overFlow
  
  Do While textPos <= textLength
    quickCalcResult = quickCalcResult + Get_NextValue(textPos, textEquation)
  Loop
  
  Unload Me
  
  Exit Sub
  
overFlow:
  MsgBox "The value caused an overflow.", vbCritical, "Overflow"
  
  txtEquation.SelStart = 0
  txtEquation.SelLength = textLength
  txtEquation.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    quickCalcResult = 0
    
    KeyCode = 0
    Shift = 0
    
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  SizeForm Me
End Sub

Private Function Get_NextValue(ByRef textPos As Long, ByRef textEquation As String) As Long
  Dim textLength As Long, started As Boolean, tVal As String, tText As String, nMode As Long
  Dim signMode As Long, tVal2 As Long
  
  textLength = Len(textEquation)
  signMode = 1
  
  Do While textPos <= textLength
    tVal = Mid$(textEquation, textPos, 1)
    tVal2 = Asc(tVal)
    
    If started Then
      If tVal = " " Or tVal = "+" Or tVal = "-" Then
        Exit Do
      ElseIf nMode = 0 Then
        If tVal2 >= 48 And tVal2 < 58 Then
          Get_NextValue = Get_NextValue * 10 + (tVal2 - 48)
        Else 'unrecognized character
          Exit Do
        End If
      ElseIf nMode = 1 Then
        If tVal2 >= 48 And tVal2 < 58 Then
          Get_NextValue = Get_NextValue * 16 + (tVal2 - 48)
        ElseIf tVal2 >= 65 And tVal2 < 71 Then
          Get_NextValue = Get_NextValue * 16 + (tVal2 - 55)
        ElseIf tVal2 >= 97 And tVal2 < 103 Then
          Get_NextValue = Get_NextValue * 16 + (tVal2 - 87)
        Else 'unrecognized character
          Exit Do
        End If
      Else
        If tVal2 >= 48 And tVal2 < 50 Then
          Get_NextValue = Get_NextValue * 2 + (tVal2 - 48)
        Else 'unrecognized character
          Exit Do
        End If
      End If
    Else
      If tVal = "-" Then
        signMode = -1
      ElseIf tVal = "H" Then
        nMode = 1
        
        started = True
      ElseIf tVal = "B" Then
        nMode = 2
        
        started = True
      ElseIf tVal2 >= 48 And tVal2 < 58 Then
        Get_NextValue = tVal2 - 48
        
        started = True
      End If
    End If
    
    textPos = textPos + 1
  Loop
  
  Get_NextValue = Get_NextValue * signMode
End Function
