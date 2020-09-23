VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmView 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9630
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
   LockControls    =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvDetails 
      Height          =   6045
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   10663
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   21167
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4260
      TabIndex        =   0
      Top             =   6270
      Width           =   1065
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FileSystem As New FileSystemObject

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim m_File As File, m_ListItem As ListItem, cString As String, tString As String, aString As String
  Dim kList() As String, kCount As Long, sList() As String, sCount As Long, lList() As String, lCount As Long
  Dim bList() As String, bCount As Long, loop1 As Long
  
  On Error Resume Next
  
  SizeForm Me
  
  Set m_File = m_FileSystem.GetFile(vFilePath & vFileName)
  
  If m_File Is Nothing Then
    Set m_ListItem = lsvDetails.ListItems.Add(, , "")
    
    With m_ListItem
      .SubItems(1) = "File details not available."
    End With
  Else
    With m_File
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Name")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = .Name
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Path")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Clean_Path(.Path)
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Extension")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Get_FileExtension
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Short Name")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = .ShortName
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Short Path")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Clean_Path(.ShortPath)
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Size")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Format(.Size, "#,###,###,##0 bytes")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "Created")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Format(.DateCreated, "dddd mmmm dd, yyyy")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "Last Modified")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Format(.DateLastModified, "dddd mmmm dd, yyyy")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "Last Accessed")
      m_ListItem.Bold = True
      m_ListItem.SubItems(1) = Format(.DateLastAccessed, "dddd mmmm dd, yyyy")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , "File Attributes")
      m_ListItem.Bold = True
      
      tString = ""
      
      m_ListItem.SubItems(1) = Format(.DateLastAccessed, "dddd mmmm dd, yyyy")
      If frmMain.chkArchive.Value = 1 Then tString = "Archive  "
      If frmMain.chkRead.Value = 1 Then tString = tString & "ReadOnly  "
      If frmMain.chkSystem.Value = 1 Then tString = tString & "System  "
      If frmMain.chkHidden.Value = 1 Then tString = tString & "Hidden  "
      
      If tString = "" Then tString = "Normal"
      m_ListItem.SubItems(1) = tString
      
      Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
      
      cString = modRegistry.Get_RegistryStringValue(HKEY_CLASSES_ROOT, Get_FileExtension, "")
      
      If cString = "" Then
        Set m_ListItem = lsvDetails.ListItems.Add(, , "Registered Type")
        m_ListItem.Bold = True
        m_ListItem.SubItems(1) = "Not registered"
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , "Registered Class")
        m_ListItem.Bold = True
        m_ListItem.SubItems(1) = "Not registered"
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , "Registered Action(s)")
        m_ListItem.Bold = True
        m_ListItem.SubItems(1) = "Registered Application(s)"
        m_ListItem.ListSubItems(1).Bold = True
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , "None")
        m_ListItem.SubItems(1) = "None"
      Else
        tString = modRegistry.Get_RegistryStringValue(HKEY_CLASSES_ROOT, cString, "")
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , "Registered Type")
        m_ListItem.Bold = True
        m_ListItem.SubItems(1) = tString
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , "Registered Class")
        m_ListItem.Bold = True
        m_ListItem.SubItems(1) = cString
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , " ")
        
        Set m_ListItem = lsvDetails.ListItems.Add(, , "Registered Action(s)")
        m_ListItem.Bold = True
        m_ListItem.SubItems(1) = "Registered Application(s)"
        m_ListItem.ListSubItems(1).Bold = True
        
        modRegistry.Get_RegistrySection HKEY_CLASSES_ROOT, cString & "\shell", kList, kCount, sList, sCount, lList, lCount, bList, bCount
        
        If kCount = 0 Then
          Set m_ListItem = lsvDetails.ListItems.Add(, , "None")
          m_ListItem.SubItems(1) = "None"
        Else
          For loop1 = 1 To kCount
            Set m_ListItem = lsvDetails.ListItems.Add(, , kList(loop1))
            
            m_ListItem.SubItems(1) = Clean_String(modRegistry.Get_RegistryStringValue(HKEY_CLASSES_ROOT, cString & "\shell\" & kList(loop1) & "\command", ""))
          Next loop1
        End If
      End If
      
      'possible future info
    End With
  End If
End Sub

Private Function Get_FileExtension()
  Dim loop1 As Long, maxLen As Long
  
  maxLen = Len(vFileName)
  
  For loop1 = maxLen To 1 Step -1
    Get_FileExtension = Mid$(vFileName, loop1, 1) & Get_FileExtension
    
    If Mid$(vFileName, loop1, 1) = "." Then Exit For
  Next loop1
End Function

Private Function Clean_String(oldString As String) As String
  Dim loop1 As Long, maxLen As Long, qCount As Long, tString As String, aString As String
  
  maxLen = Len(oldString)
  
  For loop1 = 1 To maxLen
    aString = Mid$(oldString, loop1, 1)
    
    If aString = """" Then
      If qCount = 1 Then Exit For
      
      qCount = 1
    Else
      Clean_String = Clean_String & aString
    End If
  Next loop1
  
  tString = Clean_String
  Clean_String = ""
  
  maxLen = Len(tString)
  
  For loop1 = 1 To maxLen
    aString = Mid$(tString, loop1, 1)
    
    If aString = "%" Or aString = "/" Or aString = "-" Or aString = "[" Then
      If loop1 = 1 Then Clean_String = "file is self executable or scripting"
      
      Exit For
    End If
      
    Clean_String = Clean_String & aString
  Next loop1
End Function

Private Function Clean_Path(oldString As String) As String
  Dim loop1 As Long, maxLen As Long
  
  maxLen = Len(oldString)
  
  For loop1 = maxLen To 1 Step -1
    If Mid$(oldString, loop1, 1) = "\" Then
      maxLen = loop1 - 1
      
      Exit For
    End If
  Next loop1
  
  If maxLen > 3 Then
    Clean_Path = Left$(oldString, maxLen)
  Else
    Clean_Path = Left$(oldString, 3)
  End If
End Function
