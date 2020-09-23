Attribute VB_Name = "Module1"
Option Explicit

Public prjLineHeight As Long, prjLineWidth As Long, prjBlockSize As Long

Public vFileName As String, vFilePath As String, vFileSize As Long
Public BaseOffset As Long, PageOffset As Long
Public RowOffset As Long, RowSubOffset As Boolean, LineOffset As Long
Public MaxLine As Long, LastRowLength As Long, BlockSize As Long
Public FileChange As Boolean, FileData() As Byte, LastVScroll As Long
Public FileLock As Boolean, SearchOK As Boolean, noSound As Long

Public eMode As Boolean, oldeMode As Boolean, aMode As Boolean, aMode2 As Boolean

Public SearchString As String, SearchMode As Long, SearchStart As Long, LockOut As Boolean
Public ReplaceString As String, SourceValue As Variant, RangeType As Long, quickCalcResult As Long

Public Sub SizeForm(m_Form As Form)
  Dim ratioX As Double, ratioY As Double, m_Control As Control, loop1 As Long, loop2 As Long
  Dim basetwipsX As Long, basetwipsY As Long, ratioL As Double
  
  On Error Resume Next
  
  ratioX = Screen.Width / 15360
  ratioY = Screen.Height / 11520
  
  If ratioX < ratioY Then
    ratioL = ratioX
  Else
    ratioL = ratioY
  End If
  
  basetwipsX = Screen.TwipsPerPixelX
  basetwipsY = Screen.TwipsPerPixelY
  
  If m_Form Is frmMain Then
    prjLineWidth = 16
    
    If Screen.Height <= (600 * basetwipsY) Then
      prjLineHeight = 19
    Else
      prjLineHeight = 20
    End If
    
    prjBlockSize = prjLineHeight * prjLineWidth
    
    With frmMain.Explorer1
      .Top = .Top * ratioY
      .Left = .Left * ratioX
      .Width = .Width * ratioX
      .Height = .Height * ratioY
    End With
    
    With frmMain.flxOffset
      .Top = .Top * ratioY
      .Left = .Left * ratioX
      .Width = .Width * ratioX
      .Height = .Height * ratioY
      .Font.Size = .Font.Size * ratioL
      
      .TextMatrix(0, 0) = "Address"
      .ColWidth(0) = 1100 * ratioX
      .ColAlignment(0) = 4
      
      For loop1 = 0 To prjLineHeight
        .RowHeight(loop1) = 280 * ratioY
      Next loop1
    End With
    
    With frmMain.flxHex
      .Top = .Top * ratioY
      .Left = .Left * ratioX
      .Width = .Width * ratioX
      .Height = .Height * ratioY
      .Font.Size = .Font.Size * ratioL
      
      For loop1 = 0 To prjLineWidth - 1
        .ColAlignment(loop1 * 2) = 6
        .TextMatrix(0, loop1 * 2) = MakeHexShortH(loop1)
        .ColWidth(loop1 * 2) = 280 * ratioX
        .ColAlignment(loop1 * 2 + 1) = 0
        .TextMatrix(0, loop1 * 2 + 1) = MakeHexShortL(loop1)
        .ColWidth(loop1 * 2 + 1) = 270 * ratioX
      Next loop1
      
      For loop1 = 0 To prjLineHeight
        .RowHeight(loop1) = 280 * ratioY
      Next loop1
    End With
    
    With frmMain.flxASC
      .Top = .Top * ratioY
      .Left = .Left * ratioX
      .Width = .Width * ratioX
      .Height = .Height * ratioY
      .Font.Size = .Font.Size * ratioL
      
      For loop1 = 0 To prjLineWidth - 1
        .ColWidth(loop1) = 280 * ratioX
        .ColAlignment(loop1) = 4
      Next loop1
      
      .TextMatrix(0, 3) = "A"
      .TextMatrix(0, 4) = "s"
      .TextMatrix(0, 5) = "c"
      .TextMatrix(0, 6) = "I"
      .TextMatrix(0, 7) = "I"
      .TextMatrix(0, 9) = "D"
      .TextMatrix(0, 10) = "u"
      .TextMatrix(0, 11) = "m"
      .TextMatrix(0, 12) = "p"
      
      For loop1 = 0 To prjLineHeight
        .RowHeight(loop1) = 280 * ratioY
      Next loop1
    End With
    
    For loop1 = 0 To 5
      With frmMain.Line1(loop1)
        .X2 = .X2 * ratioX
        .X1 = .X1 * ratioX
        .Y2 = .Y2 * ratioY
        .Y1 = .Y1 * ratioY
      End With
    Next loop1
    
    'm_Form.DrawWidth = 2
    'm_Form.Line (basetwipsX, basetwipsY)-(Screen.Width - 2 * basetwipsX, Screen.Height - basetwipsY * 2), &HFFFFFF, B
    'm_Form.DrawWidth = 1
    'm_Form.Line (0, 0)-(Screen.Width - basetwipsX, Screen.Height - basetwipsY), 0, B
    'm_Form.Line (3 * basetwipsX, 3 * basetwipsY)-(Screen.Width - 4 * basetwipsX, Screen.Height - 4 * basetwipsY), 0, B
  Else
    With m_Form
      .Width = .Width * ratioX
      .Height = .Height * ratioX
    End With
    
    m_Form.DrawWidth = 2
    m_Form.Line (basetwipsX, basetwipsY)-(m_Form.Width - 2 * basetwipsX, m_Form.Height - basetwipsY * 2), &HFFFFFF, B
    m_Form.DrawWidth = 1
    m_Form.Line (0, 0)-(m_Form.Width - basetwipsX, m_Form.Height - basetwipsY), 0, B
    m_Form.Line (3 * basetwipsX, 3 * basetwipsY)-(m_Form.Width - 4 * basetwipsX, m_Form.Height - 4 * basetwipsY), 0, B
  End If
  
  For Each m_Control In m_Form.Controls
    If TypeOf m_Control Is CommandButton Or TypeOf m_Control Is Label Or TypeOf m_Control Is OptionButton Or TypeOf m_Control Is TextBox Or TypeOf m_Control Is Image Or TypeOf m_Control Is Frame Or TypeOf m_Control Is CheckBox Or TypeOf m_Control Is VScrollBar Or TypeOf m_Control Is ComboBox Then
      With m_Control
        .Top = .Top * ratioY
        .Left = .Left * ratioX
        .Width = .Width * ratioX
        If Not (TypeOf m_Control Is ComboBox) Then .Height = .Height * ratioY
        
        If Not (TypeOf m_Control Is Image Or TypeOf m_Control Is VScrollBar) Then .Font.Size = .Font.Size * ratioL
      End With
    ElseIf TypeOf m_Control Is ListView Then
      With m_Control
        .Top = .Top * ratioY
        .Left = .Left * ratioX
        .Width = .Width * ratioX
        .Height = .Height * ratioY
        
        .Font.Size = .Font.Size * ratioL
        
        .ColumnHeaders(1).Width = .ColumnHeaders(1).Width * ratioX
        .ColumnHeaders(2).Width = .ColumnHeaders(2).Width * ratioX
      End With
    End If
  Next m_Control
End Sub

Public Function MakeHexLong(ByVal vall As Long) As String
  Dim valh As Long
  
  valh = vall \ 268435456
  vall = vall Mod 268435456
  
  If valh > 9 Then
    MakeHexLong = Chr$(55 + valh)
  Else
    MakeHexLong = Chr$(48 + valh)
  End If
  
  valh = vall \ 16777216
  vall = vall Mod 16777216
    
  If valh > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + valh)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + valh)
  End If
  
  valh = vall \ 1048576
  vall = vall Mod 1048576
    
  If valh > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + valh)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + valh)
  End If
  
  valh = vall \ 65536
  vall = vall Mod 65536
    
  If valh > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + valh)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + valh)
  End If
  
  valh = vall \ 4096
  vall = vall Mod 4096
    
  If valh > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + valh)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + valh)
  End If
  
  valh = vall \ 256
  vall = vall Mod 256
    
  If valh > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + valh)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + valh)
  End If
  
  valh = vall \ 16
  vall = vall Mod 16
    
  If valh > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + valh)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + valh)
  End If
  
  If vall > 9 Then
    MakeHexLong = MakeHexLong & Chr$(55 + vall)
  Else
    MakeHexLong = MakeHexLong & Chr$(48 + vall)
  End If
End Function

Public Function MakeHexShort(ByVal vall As Long) As String
  Dim valh As Long
    
  valh = vall \ 16
  vall = vall Mod 16
  
  If valh > 9 Then
    MakeHexShort = Chr$(55 + valh)
  Else
    MakeHexShort = Chr$(48 + valh)
  End If
  
  If vall > 9 Then
    MakeHexShort = MakeHexShort & Chr$(55 + vall)
  Else
    MakeHexShort = MakeHexShort & Chr$(48 + vall)
  End If
End Function

Public Function MakeHexShortL(ByVal vall As Long) As String
  vall = vall Mod 16
  
  If vall > 9 Then
    MakeHexShortL = Chr$(55 + vall)
  Else
    MakeHexShortL = Chr$(48 + vall)
  End If
End Function

Public Function MakeHexShortH(ByVal valh As Long) As String
  valh = valh \ 16
  
  If valh > 9 Then
    MakeHexShortH = Chr$(55 + valh)
  Else
    MakeHexShortH = Chr$(48 + valh)
  End If
End Function
