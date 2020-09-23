VERSION 5.00
Begin VB.Form frmDelete 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuantity 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   900
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   180
      TabIndex        =   1
      Top             =   1350
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2040
      TabIndex        =   2
      Top             =   1350
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "starting from (and including) cursor."
      Height          =   585
      Left            =   210
      TabIndex        =   5
      Top             =   690
      Width           =   2985
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "byte(s)"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   300
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Delete"
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   300
      Width           =   630
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim qVal As Long, cPos As Long
  Dim bSize As Long, bPos As Long, bData() As Byte
  
  On Error Resume Next
  
  qVal = val(txtQuantity.Text)
  cPos = BaseOffset + PageOffset + (LineOffset * prjLineWidth) + RowOffset + 1
  
  If qVal <= 0 Or cPos + qVal > vFileSize + 1 Then
    MsgBox "The quantity is invalid!", vbCritical, "Invalid Value"
    
    With txtQuantity
      .SetFocus
      
      .SelStart = 0
      .SelLength = Len(.Text)
    End With
  ElseIf MsgBox("Warning! This action cannot be undone. Are you sure you wish to proceed?", vbYesNo, "Delete Block") = vbYes Then
    Kill vFilePath & "tempFileAdjust.temp"
    
    Open vFilePath & "tempFileAdjust.temp" For Binary As #2
    
    bPos = 1
    
    Do While bPos < cPos
      If cPos - bPos > 50000 Then
        bSize = 50000
      Else
        bSize = cPos - bPos
      End If
      
      ReDim bData(1 To bSize)
      
      Get #1, bPos, bData
      Put #2, bPos, bData
      
      bPos = bPos + bSize
    Loop
    
    Do While bPos < vFileSize - qVal + 1
      If vFileSize + 1 - qVal - bPos > 50000 Then
        bSize = 50000
      Else
        bSize = vFileSize + 1 - qVal - bPos
      End If
      
      ReDim bData(1 To bSize)
      
      Get #1, bPos + qVal, bData
      Put #2, bPos, bData
      
      bPos = bPos + bSize
    Loop
    
    Close #1
    Close #2
    
    Kill vFilePath & vFileName
    Name vFilePath & "tempFileAdjust.temp" As vFilePath & vFileName
    
    Open vFilePath & vFileName For Binary As #1
    
    vFileSize = vFileSize - qVal
    frmMain.lblFileStats(1).Caption = Format(vFileSize, "##,###,###,##0") & " Bytes"
    
    frmMain.Explorer1.Refresh True, -qVal
    
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  SizeForm Me
  
  txtQuantity.SelLength = 1
End Sub
