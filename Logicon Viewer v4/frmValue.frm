VERSION 5.00
Begin VB.Form frmValue 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   75
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   2010
      Width           =   855
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter"
      Default         =   -1  'True
      Height          =   375
      Left            =   2610
      TabIndex        =   2
      Top             =   2010
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   3285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1065
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   3285
   End
End
Attribute VB_Name = "frmValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnter_Click()
  Dim NewValue As Long
  
  On Error GoTo badresult
  
  NewValue = CLng(Text1.Text)
  
  Select Case RangeType
    Case 0:
      If NewValue >= 0 And NewValue <= 255 Then
        SourceValue = NewValue
        
        Unload Me
        
        Exit Sub
      End If
    Case 1:
      If NewValue >= -128 And NewValue <= 127 Then
        SourceValue = NewValue
        
        Unload Me
        
        Exit Sub
      End If
    Case 2:
      If NewValue >= 0 And NewValue <= 65535 Then
        SourceValue = NewValue
        
        Unload Me
        
        Exit Sub
      End If
    Case 3:
      If NewValue >= -32768 And NewValue <= 32767 Then
        SourceValue = NewValue
        
        Unload Me
        
        Exit Sub
      End If
    Case 4:
      If NewValue >= -2147483648# And NewValue <= 2147483647 Then
        SourceValue = NewValue
        
        Unload Me
        
        Exit Sub
      End If
  End Select
  
badresult:
  MsgBox "The value was invalid!", vbOKOnly
  
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  Text1.SetFocus
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  SizeForm Me
  
  Select Case RangeType
    Case 0:
      Me.Caption = "Unsigned Byte"
      Label1.Caption = "Value must be from" & vbNewLine & "0" & vbNewLine & "to" & vbNewLine & "255"
    Case 1:
      Me.Caption = "Signed Byte"
      Label1.Caption = "Value must be from" & vbNewLine & "-128" & vbNewLine & "to" & vbNewLine & "+127"
    Case 2:
      Me.Caption = "Unsigned Double Byte"
      Label1.Caption = "Value must be from" & vbNewLine & "0" & vbNewLine & "to" & vbNewLine & "65,535"
    Case 3:
      Me.Caption = "Signed Double Byte"
      Label1.Caption = "Value must be from" & vbNewLine & "-32,768" & vbNewLine & "to" & vbNewLine & "+32,767"
    Case 4:
      Me.Caption = "Signed Quad Byte"
      Label1.Caption = "Value must be from" & vbNewLine & "-2,147,483,648" & vbNewLine & "to" & vbNewLine & "+2,147,483,647"
  End Select
  
  If SourceValue >= 0 Then
    If RangeType = 0 Or RangeType = 2 Then
      Text1.Text = Format(SourceValue, "#,###,###,##0")
    Else
      Text1.Text = Format(SourceValue, "+#,###,###,##0")
    End If
  Else
    Text1.Text = Format(SourceValue, "#,###,###,##0")
  End If
  
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
  
  SourceValue = "Cancel"
End Sub

