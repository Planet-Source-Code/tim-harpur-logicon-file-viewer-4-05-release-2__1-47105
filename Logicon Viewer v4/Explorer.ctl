VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Explorer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   LockControls    =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   11190
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Refresh directory structure."
      Top             =   4200
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   210
      Top             =   3840
   End
   Begin VB.Timer tmrVerify 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3870
   End
   Begin MSComctlLib.ListView lsvFileList 
      Height          =   3750
      Left            =   4800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   6615
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList2"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size (Bytes)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSmallTreeIcons 
      Left            =   5070
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":0228
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":033C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":0450
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.ctx":09B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3750
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6615
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgSmallTreeIcons"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFilePathName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   4
      Top             =   4170
      Width           =   11085
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4815
      TabIndex        =   3
      Top             =   3810
      Width           =   6315
   End
   Begin VB.Label lblDirStats 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Top             =   3810
      Width           =   4725
   End
End
Attribute VB_Name = "Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FileSystem As New FileSystemObject, m_LastNode As Node

Private DriveSerial(1 To 26) As Double, DriveType(1 To 26) As Long
Private DriveHomeNode(1 To 26) As Node, MasterNode As Node
Private LockOut As Boolean, NoAction As Long

Private m_FileName As String, m_FilePath As String, m_FilePathEx As String
Private m_FileSize As Long
Private m_FileAttributes As Long, m_FileError As Long
Private m_Count As Long, d_Count As Long, m_Size As Long

Private m_FileFilter As String, m_LabelEdit As Boolean

Public Event FileSelected()
Public Event FileLaunched()
Public Event ClickSound()
Public Event FileEditStart(oldPathName As String)
Public Event PathRenamed(oldPath As String, newPath As String)
Public Event FileRenamed(oldFilePathName As String, newFilePath As String, newFileName As String)

Public Property Get labelEditStatus() As Boolean
  labelEditStatus = m_LabelEdit
End Property

Public Property Get FilePathName() As String
  FilePathName = m_FilePathEx & m_FileName
End Property

Public Property Get FileType() As String
  Dim m_File As File
  
  On Error GoTo nofile
  
  If m_FileName <> "" Then
    Set m_File = m_FileSystem.GetFile(m_FilePathEx & m_FileName)
    
    FileType = m_File.Type
    
    Exit Sub
  End If
  
nofile:
  FileType = ""
End Property

Public Property Let FilePathName(FileName As String)
  Dim displayed_path As String, old_path As String
  
  On Error Resume Next
  
  displayed_path = FixPath(m_LastNode.FullPath)
  old_path = m_FilePath
  
  m_FilePath = m_FileSystem.GetParentFolderName(FileName)
  
  If Right$(m_FilePath, 1) <> "\" Then
    m_FilePathEx = m_FilePath & "\"
  Else
    m_FilePathEx = m_FilePath
  End If
  
  m_FileName = m_FileSystem.GetFileName(FileName)
  
  On Error GoTo badfile
  
  m_FileSize = FileLen(m_FilePathEx & m_FileName)
  m_FileAttributes = GetAttr(m_FilePathEx & m_FileName)
  m_FileError = 0
  
  lblFilePathName.Caption = m_FilePathEx & m_FileName
  
  If (displayed_path <> "" And displayed_path = old_path) Or displayed_path = m_FilePath Then Reload_FileList displayed_path
  
  Exit Property

badfile:
  m_FileName = ""
  m_FilePath = ""
  m_FilePathEx = ""
  m_FileSize = 0
  m_FileAttributes = 0
  m_FileError = 1
  
  lblFilePathName.Caption = ""
  
  If (displayed_path <> "" And displayed_path = old_path) Or displayed_path = m_FilePath Then Reload_FileList displayed_path
End Property

Public Property Let FileFilter(FileType As String)
  m_FileFilter = FileType
  
  If m_FileFilter = "*.*" Then m_FileFilter = ""
  m_FileFilter = LCase(m_FileFilter)
  
  If Not (m_LastNode Is Nothing) Then Reload_FileList FixPathEx(m_LastNode.FullPath)
End Property

Public Property Get FileFilter() As String
  FileFilter = m_FileFilter
End Property

Public Property Get FileName() As String
  FileName = m_FileName
End Property

Public Property Get FilePath() As String
  FilePath = m_FilePath
End Property

Public Property Get FileSize() As Long
  FileSize = m_FileSize
End Property

Public Property Get FileAttributes() As Long
Attribute FileAttributes.VB_Description = "Reports when the user has selected a new file in the filelist box, or when the current file has become invalid from such as disk being removed."
Attribute FileAttributes.VB_MemberFlags = "200"
  FileAttributes = m_FileAttributes
End Property

Public Property Get FileError() As Long
  FileError = m_FileError
End Property

Public Sub flvPicture(pictureObject As Picture)
  lsvFileList.Picture = pictureObject
End Sub

Private Sub cmdRefresh_Click()
  Refresh_DirectoryTree
End Sub

Private Sub lsvFileList_AfterLabelEdit(Cancel As Integer, NewString As String)
  Dim oldPathName As String, FilePath As String
  
  m_LabelEdit = False
  
  oldPathName = m_FilePathEx & m_FileName
  
  'give user program a chance to release the file handle if this file is already open
  RaiseEvent FileEditStart(oldPathName)
  
  On Error GoTo cantRename
  Name oldPathName As m_FilePathEx & NewString
  
  RaiseEvent FileRenamed(oldPathName, m_FilePath, NewString)
  
  m_FileName = NewString
  lblFilePathName.Caption = m_FilePathEx & m_FileName
  
  Exit Sub

cantRename:
  Cancel = 1
End Sub

Private Sub lsvFileList_BeforeLabelEdit(Cancel As Integer)
  m_LabelEdit = True
End Sub

Private Sub lsvFileList_Click()
  Dim nodeName As String, tNode As Node
  
  m_LabelEdit = False
  
  If LockOut Then Exit Sub
  
  If lsvFileList.ListItems.Count < 1 Then Exit Sub
  
  If m_FileName = lsvFileList.SelectedItem.Text And m_FilePath = FixPath(TreeView1.SelectedItem.FullPath) Then Exit Sub
  
  LockOut = True
  
  If lsvFileList.SelectedItem.ListSubItems(1).Text = "sub-Directory" Then
    On Error Resume Next
    
    nodeName = lsvFileList.SelectedItem.Text
    nodeName = Right$(nodeName, Len(nodeName) - 4)
    
    m_LastNode.Expanded = True
    
    Set tNode = m_LastNode.Child
    
    Do While Not (tNode Is Nothing)
      If tNode.Text = nodeName Then
        tNode.Selected = True
        
        Set m_LastNode = tNode
        
        Reload_FileList FixPath(tNode.FullPath)
        
        Exit Do
      Else
        Set tNode = tNode.Next
      End If
    Loop
  Else
    On Error GoTo badfile
    
    m_FileName = lsvFileList.SelectedItem.Text
    m_FilePath = FixPath(TreeView1.SelectedItem.FullPath)
    m_FilePathEx = FixPathEx(TreeView1.SelectedItem.FullPath)
    
    m_FileSize = FileLen(m_FilePathEx & m_FileName)
    m_FileAttributes = GetAttr(m_FilePathEx & m_FileName)
    m_FileError = 0
    
    lblFilePathName.Caption = m_FilePathEx & m_FileName
    
    RaiseEvent ClickSound
    RaiseEvent FileSelected
  End If
  
  LockOut = False
  
  Exit Sub
  
badfile:
  m_FileName = ""
  m_FilePath = ""
  m_FilePathEx = ""
  m_FileSize = 0
  m_FileAttributes = 0
  m_FileError = 1
  
  lblFilePathName.Caption = ""
  
  RaiseEvent FileSelected
  
  LockOut = False
End Sub

Private Sub lsvFileList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim m_position As Long
  
  m_LabelEdit = False
  
  RaiseEvent ClickSound
  
  m_position = ColumnHeader.position - 1
  
  If m_position = 1 Then m_position = 2
  
  If lsvFileList.SortKey = m_position Then
    If lsvFileList.SortOrder = lvwAscending Then
      lsvFileList.SortOrder = lvwDescending
    Else
      lsvFileList.SortOrder = lvwAscending
    End If
  Else
    lsvFileList.SortKey = m_position
    lsvFileList.SortOrder = lvwAscending
  End If
End Sub

Private Sub lsvFileList_DblClick()
  RaiseEvent FileLaunched
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  
  'initialize this control
  
  NoAction = 0
  LockOut = False
  Refresh_DirectoryTree
End Sub

Private Sub tmrVerify_Timer()
  Dim loop1 As Long, m_drive As Drive, vName As String
  
  If NoAction > 0 Then
    NoAction = NoAction - 1
    
    If NoAction = 0 Then
      TreeView1.Enabled = True
    End If
  End If
  
  'check cd/dvd drives and update if nescesary
  
  For loop1 = 1 To 26
    If DriveType(loop1) = 4 Then
      Set m_drive = m_FileSystem.GetDrive(Chr$(loop1 + 64) & ":")
      
      With m_drive
        If .IsReady Then
          If DriveSerial(loop1) <> .SerialNumber Then
            TreeView1.Nodes.Remove DriveHomeNode(loop1).Index
            
            DriveSerial(loop1) = .SerialNumber
            
            vName = .VolumeName
            If Len(vName) < 11 Then
              vName = "(" & vName & ")" & Space(11 - Len(vName))
            ElseIf Len(vName) > 11 Then
              vName = "(" & Left$(vName, 11) & ")"
            Else
              vName = "(" & vName & ")"
            End If
            
            Set DriveHomeNode(loop1) = TreeView1.Nodes.Add(, , , Chr$(64 + loop1) & ": " & vName, 3)
            DriveHomeNode(loop1).Sorted = True
            DriveHomeNode(loop1).Bold = True
            
            GenerateSubBranchX DriveHomeNode(loop1), Left$(DriveHomeNode(loop1).Text, 2) & "\"
            
            DriveHomeNode(loop1).Tag = 1
            DriveHomeNode(loop1).ForeColor = 0
          End If
        ElseIf DriveSerial(loop1) <> 0 Then
          'mark this node as removable and clear
          
          If Not (m_LastNode Is Nothing) Then
            If Left$(m_LastNode.FullPath, 2) = Left$(DriveHomeNode(loop1).FullPath, 2) Then
              lsvFileList.ListItems.Clear
              
              Set m_LastNode = MasterNode
            End If
          End If
          
          If Left$(m_FilePath, 2) = Left$(DriveHomeNode(loop1).FullPath, 2) Then
            m_FileName = ""
            m_FilePath = ""
            m_FileSize = 0
            m_FileAttributes = 0
            m_FileError = 2
            
            lblFilePathName.Caption = ""
            
            'signal that disk has been removed
            
            RaiseEvent FileSelected
          End If
          
          TreeView1.Nodes.Remove DriveHomeNode(loop1).Index
          
          DriveSerial(loop1) = 0
          
          Set DriveHomeNode(loop1) = TreeView1.Nodes.Add(, , , Chr$(64 + loop1) & ": (no disk present)", 3)
          DriveHomeNode(loop1).Sorted = True
          DriveHomeNode(loop1).Bold = True
          
          DriveHomeNode(loop1).Tag = "Removable"
          DriveHomeNode(loop1).ForeColor = &HA0A0A0
        End If
      End With
    End If
  Next loop1
  
  If Not (TreeView1.SelectedItem Is m_LastNode) Then
    Set TreeView1.SelectedItem = m_LastNode
  End If
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
  Dim oldName As String, newName As String
  Dim oldPath As String, loop1 As Long
  
  m_LabelEdit = False
  
  oldPath = FixPath(m_LastNode.FullPath)
  
  loop1 = Len(oldPath) - 1
  
  Do While loop1 > 0
    If Mid$(oldPath, loop1, 1) = "\" Then Exit Do
    
    loop1 = loop1 - 1
  Loop
  
  oldPath = Left$(oldPath, loop1)
  
  oldName = oldPath & m_LastNode.Text
  newName = oldPath & NewString
  
  On Error GoTo cantRename
  Name oldName As newName
  
  m_FilePath = newName
  m_FilePathEx = newName & "\"
  
  lblFilePathName.Caption = m_FilePathEx & m_FileName
  
  RaiseEvent PathRenamed(oldName, newName)
  
  Exit Sub

cantRename:
  Cancel = 1
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
  If m_LastNode.parent Is Nothing Then Cancel = 1
  
  m_LabelEdit = True
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
  m_LabelEdit = False
  
  On Error Resume Next
  
  If NoAction > 0 Then Exit Sub
  NoAction = 3
  
  tmrVerify.Enabled = False
  TreeView1.Enabled = False
  
  If CheckDrive(Node) Then
    tmrVerify.Enabled = True
    
    Exit Sub
  End If
  
  If Node.Tag = "Removable" Then
    tmrVerify.Enabled = True
    
    Exit Sub
  End If
  
  RaiseEvent ClickSound
    
  Reload_FileList FixPath(Node.FullPath)
  
  Set m_LastNode = Node
  
  tmrVerify.Enabled = True
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
  On Error Resume Next
  
  m_LabelEdit = False
  
  If CheckDrive(Node) Then Exit Sub
  
  RaiseEvent ClickSound
  
  GenerateSubBranch Node, FixPath(Node.FullPath)
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  'select directory
  
  m_LabelEdit = False
  
  On Error Resume Next
  
  If CheckDrive(Node) Then Exit Sub
  If Node.Tag = "Removable" Then Exit Sub
  
  RaiseEvent ClickSound
  
  Set m_LastNode = Node
  
  Reload_FileList FixPath(Node.FullPath)
End Sub

Private Sub Refresh_DirectoryTree()
  Dim loop1 As Long, vName As String, drivenumber As Long
  Dim m_DriveList As Drives, m_drive As Drive
  Dim m_Node As Node, m_Nodes As Nodes
  
  tmrVerify.Enabled = False
  
  m_LabelEdit = False
  
  TreeView1.Nodes.Clear
  
  Set m_LastNode = Nothing
  Set m_DriveList = m_FileSystem.Drives
  
  For loop1 = 1 To 26
    DriveSerial(loop1) = 0
    DriveType(loop1) = -1
    Set DriveHomeNode(loop1) = Nothing
  Next loop1
  
  For Each m_drive In m_DriveList
    Select Case m_drive.DriveType
      Case 0: 'unknown drives
                  Set m_Node = TreeView1.Nodes.Add(, , , m_drive.DriveLetter & ": (unknown access not supported)", 3)
                  m_Node.Tag = "Removable"
                  m_Node.ForeColor = &HA0A0A0
                  m_Node.Bold = True
                  
                  drivenumber = Asc(UCase(m_drive.DriveLetter)) - 64
                  Set DriveHomeNode(drivenumber) = m_Node
                  DriveType(drivenumber) = 0
                  
      Case 1: 'floppy drives
                  Set m_Node = TreeView1.Nodes.Add(, , , m_drive.DriveLetter & ":", 1)
                  m_Node.Tag = "Removable"
                  m_Node.ForeColor = &HA0A0A0
                  m_Node.Bold = True
                  
                  drivenumber = Asc(UCase(m_drive.DriveLetter)) - 64
                  Set DriveHomeNode(drivenumber) = m_Node
                  DriveType(drivenumber) = 1
                  
      Case 2: 'hard drives
                  vName = m_drive.VolumeName
                  If Len(vName) < 11 Then
                    vName = "(" & vName & ")" & Space(11 - Len(vName))
                  ElseIf Len(vName) > 11 Then
                    vName = "(" & Left$(vName, 11) & ")"
                  Else
                    vName = "(" & vName & ")"
                  End If
                  
                  Set m_Node = TreeView1.Nodes.Add(, , , m_drive.DriveLetter & ": " & vName, 2)
                  m_Node.Sorted = True
                  m_Node.Tag = 1
                  m_Node.Bold = True
                  
                  GenerateSubBranchX m_Node, Left$(m_Node.Text, 2) & "\"
                  
                  If m_LastNode Is Nothing Then
                    Set m_LastNode = m_Node
                    Set MasterNode = m_Node
                  End If
                  
                  drivenumber = Asc(UCase(m_drive.DriveLetter)) - 64
                  Set DriveHomeNode(drivenumber) = m_Node
                  DriveType(drivenumber) = 2
                  DriveSerial(drivenumber) = m_drive.SerialNumber
                  
      Case 3: 'network drives
                  Set m_Node = TreeView1.Nodes.Add(, , , m_drive.DriveLetter & ": (network drive)", 7)
                  m_Node.Tag = "Removable"
                  m_Node.ForeColor = &HA0A0A0
                  m_Node.Bold = True
                  
                  drivenumber = Asc(UCase(m_drive.DriveLetter)) - 64
                  Set DriveHomeNode(drivenumber) = m_Node
                  DriveType(drivenumber) = 3
                  
      Case 4: 'cd/dvd drives
                  Set m_Node = TreeView1.Nodes.Add(, , , m_drive.DriveLetter & ": (no disk present)", 3)
                  m_Node.Tag = "Removable"
                  m_Node.ForeColor = &HA0A0A0
                  m_Node.Bold = True
                  
                  drivenumber = Asc(UCase(m_drive.DriveLetter)) - 64
                  Set DriveHomeNode(drivenumber) = m_Node
                  DriveType(drivenumber) = 4
                  
      Case 5: 'ram-drive
                  vName = m_drive.VolumeName
                  If Len(vName) < 11 Then
                    vName = "(" & vName & ")" & Space(11 - Len(vName))
                  ElseIf Len(vName) > 11 Then
                    vName = "(" & Left$(vName, 11) & ")"
                  Else
                    vName = "(" & vName & ")"
                  End If
                  
                  Set m_Node = TreeView1.Nodes.Add(, , , m_drive.DriveLetter & ": " & vName, 7)
                  m_Node.Sorted = True
                  m_Node.Tag = 1
                  m_Node.Bold = True
                  
                  GenerateSubBranchX m_Node, Left$(m_Node.Text, 2) & "\"
                  
                  If m_LastNode Is Nothing Then Set m_LastNode = m_Node
                  
                  drivenumber = Asc(UCase(m_drive.DriveLetter)) - 64
                  Set DriveHomeNode(drivenumber) = m_Node
                  DriveType(drivenumber) = 5
    End Select
  Next
  
  m_LastNode.Selected = True
  TreeView1_NodeClick m_LastNode
  
  tmrVerify.Enabled = True
End Sub

Private Sub GenerateSubBranch(parent As Node, Path As String)
  Dim m_Node As Node, m_Nodes As Long, mString As String
  
  If parent.Tag = 2 Then Exit Sub
  parent.Tag = 2
  
  m_Nodes = parent.Children
  
  If m_Nodes > 0 Then
    Set m_Node = parent.Child
    
    Do While m_Nodes > 0
      With m_Node
        mString = FixPath(.FullPath)
        
        If Not (m_FileSystem.GetFolder(mString).SubFolders Is Nothing) Then GenerateSubBranchX m_Node, mString
        
        m_Nodes = m_Nodes - 1
        Set m_Node = .Next
      End With
    Loop
  End If
End Sub

Private Sub GenerateSubBranchX(parent As Node, Path As String)
  Dim m_folder As Folder, m_folders As Folders, m_Node As Node
  
  Set m_folder = m_FileSystem.GetFolder(Path)
  Set m_folders = m_folder.SubFolders
  
  For Each m_folder In m_folders
    Set m_Node = TreeView1.Nodes.Add(parent.Index, tvwChild, , m_folder.Name, 4)
    
    With m_Node
      .ExpandedImage = 5
      .Sorted = True
      .Tag = 0
    End With
  Next
End Sub

Private Sub Reload_FileList(ByVal Path As String)
  Dim m_item As ListItem, fileIsHere As Boolean
  Dim t_size As Long, ts_size As String, m_folder As Folder, m_File As File
  Dim m_subfolder As Folder, m_files As Files, m_subfolders As Folders
  
  On Error GoTo thatsit
  lsvFileList.Sorted = False
  'lsvFileList.Visible = False
  
  m_LabelEdit = False
  
  m_Count = 0
  d_Count = 0
  m_Size = 0
  
  lsvFileList.ListItems.Clear
  
  lblStats.Caption = "Scanning directory for files..."
  lblDirStats.Caption = ""
  
  DoEvents
  
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  fileIsHere = (Path = m_FilePath) And (m_FileSize > 0)
  
  Set m_folder = m_FileSystem.GetFolder(Path)
    
  With m_folder
    Set m_subfolders = .SubFolders
    Set m_files = .Files
    
    If .IsRootFolder Then
      lblDirStats.Caption = m_folder.Drive.VolumeName
    ElseIf .ParentFolder.IsRootFolder Then
      lblDirStats.Caption = .Drive.VolumeName & "\" & .Name
    Else
      lblDirStats.Caption = .Drive.VolumeName & "\ ... \" & .Name
    End If
    
    If fileIsHere Then
      For Each m_File In m_files
        With m_File
          Set m_item = lsvFileList.ListItems.Add(, , .Name)
          
          t_size = .Size
          ts_size = Format(t_size, "#,###,###,##0")
          
          With m_item
            .SubItems(1) = ts_size
            .SubItems(2) = Chr$(64 + Len(ts_size)) & ts_size
            .Selected = (m_FileName = m_File.Name)
            
            m_Size = m_Size + t_size
          End With
        End With
      Next m_File
    Else
      For Each m_File In m_files
        With m_File
          Set m_item = lsvFileList.ListItems.Add(, , .Name)
          
          t_size = .Size
          ts_size = Format(t_size, "#,###,###,##0")
          
          With m_item
            .SubItems(1) = ts_size
            .SubItems(2) = Chr$(64 + Len(ts_size)) & ts_size
            .Selected = False
            
            m_Size = m_Size + t_size
          End With
        End With
      Next m_File
    End If
    
    For Each m_subfolder In m_subfolders
      With m_subfolder
        If .Name <> "." And .Name <> ".." Then
          Set m_item = lsvFileList.ListItems.Add(, , " ..." & .Name)
          
          With m_item
            .Bold = True
            .SubItems(1) = "sub-Directory"
            .ListSubItems(1).Bold = True
            .SubItems(2) = "0"
            .Selected = False
          End With
        End If
      End With
    Next m_subfolder
    
    m_Count = m_files.Count
    d_Count = m_subfolders.Count
  End With
  
  lsvFileList.Sorted = True
  lsvFileList.ListItems(1).EnsureVisible
  'lsvFileList.Visible = True
thatsit:
  lblStats = d_Count & " sub-Directories" & " / " & m_Count & " Files @ " & Format(m_Size, "##,###,###,##0 bytes")
End Sub

Private Function FixPath(oldPath As String) As String
  If Len(oldPath) = 16 Then
    FixPath = Left$(oldPath, 2) & "\"
  Else
    FixPath = Left$(oldPath, 2) & Right$(oldPath, Len(oldPath) - 16)
  End If
End Function

Private Function FixPathEx(oldPath As String) As String
  If Len(oldPath) = 16 Then
    FixPathEx = Left$(oldPath, 2) & "\"
  Else
    FixPathEx = Left$(oldPath, 2) & Right$(oldPath, Len(oldPath) - 16) & "\"
  End If
End Function

Private Function CheckDrive(c_Node As Node) As Boolean
  Dim m_drive As Drive, vName As String, driveNum As Long
  
  'check if drive has been removed or changed and update if nescesary
  
  CheckDrive = False
  
  driveNum = Asc(UCase(Left$(c_Node.FullPath, 1))) - 64
  
  Set m_drive = m_FileSystem.GetDrive(Chr$(driveNum + 64) & ":")
    
  With m_drive
    If .IsReady Then
      If DriveSerial(driveNum) <> .SerialNumber Then
        CheckDrive = True
        
        TreeView1.Nodes.Remove DriveHomeNode(driveNum).Index
        
        DriveSerial(driveNum) = .SerialNumber
        
        vName = .VolumeName
        If Len(vName) < 11 Then
          vName = "(" & vName & ")" & Space(11 - Len(vName))
        ElseIf Len(vName) > 11 Then
          vName = "(" & Left$(vName, 11) & ")"
        Else
          vName = "(" & vName & ")"
        End If
        
        Set DriveHomeNode(driveNum) = TreeView1.Nodes.Add(, , , Chr$(64 + driveNum) & ": " & vName, 3)
        DriveHomeNode(driveNum).Sorted = True
        DriveHomeNode(driveNum).Bold = True
        
        GenerateSubBranchX DriveHomeNode(driveNum), Left$(DriveHomeNode(driveNum).Text, 2) & "\"
        
        DriveHomeNode(driveNum).Tag = 1
        DriveHomeNode(driveNum).ForeColor = 0
        
        Set m_LastNode = DriveHomeNode(driveNum)
        Reload_FileList FixPath(m_LastNode.FullPath)
      End If
    ElseIf DriveSerial(driveNum) <> 0 Then
      'mark this node as removable and clear
      
      CheckDrive = True
      
      If Not (m_LastNode Is Nothing) Then
        If Left$(m_LastNode.FullPath, 2) = Left$(DriveHomeNode(driveNum).FullPath, 2) Then
          lsvFileList.ListItems.Clear
          
          Set m_LastNode = MasterNode
        End If
      End If
      
      If Left$(m_FilePath, 2) = Left$(DriveHomeNode(driveNum).FullPath, 2) Then
        m_FileName = ""
        m_FilePath = ""
        m_FileSize = 0
        m_FileAttributes = 0
        m_FileError = 2
        
        lblFilePathName.Caption = ""
        
        'signal that disk has been removed
        
        RaiseEvent FileSelected
      End If
        
      TreeView1.Nodes.Remove DriveHomeNode(driveNum).Index
        
      DriveSerial(driveNum) = 0
        
      Set DriveHomeNode(driveNum) = TreeView1.Nodes.Add(, , , Chr$(64 + driveNum) & ":", 3)
      DriveHomeNode(driveNum).Sorted = True
      DriveHomeNode(driveNum).Bold = True
        
      DriveHomeNode(driveNum).Tag = "Removable"
      DriveHomeNode(driveNum).ForeColor = &HA0A0A0
    End If
  End With
End Function

Private Sub UserControl_Initialize()
  Timer1.Enabled = True
  
  m_FileName = ""
  m_FilePath = ""
  m_FileSize = 0
  m_FileAttributes = 0
  m_FileError = 0
  m_LabelEdit = False
End Sub

Private Sub UserControl_LostFocus()
  m_LabelEdit = False
End Sub

Private Sub UserControl_Resize()
  Dim sizeAdjust_w As Double, sizeAdjust_h As Double, adjustFactor As Double
  Dim newHeaderSize As Long
  
  On Error Resume Next
  
  If UserControl.Width < 2000 Then UserControl.Width = 2000
  If UserControl.Height < 2000 Then UserControl.Height = 2000
  
  sizeAdjust_w = (UserControl.Width - 90) / (11190 - 90)
  sizeAdjust_h = (UserControl.Height - 120 - 630) / (4545 - 120 - 630)
  
  If sizeAdjust_w < 0.8 Then
    TreeView1.Font.Size = 8
    lsvFileList.Font.Size = 8
    lblDirStats.Font.Size = 8
    lblDirStats.Font.Bold = False
    lblStats.Font.Size = 8
    lblStats.Font.Bold = False
    lblFilePathName.Font.Size = 8
    lblFilePathName.Font.Bold = False
  Else
    adjustFactor = (0.5 + sizeAdjust_w / 2)
    
    TreeView1.Font.Size = 10 * adjustFactor
    lsvFileList.Font.Size = 10 * adjustFactor
    lblDirStats.Font.Size = 10 * adjustFactor
    lblDirStats.Font.Bold = True
    lblStats.Font.Size = 10 * adjustFactor
    lblStats.Font.Bold = True
    lblFilePathName.Font.Size = 10 * adjustFactor
    lblFilePathName.Font.Bold = True
  End If
  
  TreeView1.Left = 30
  TreeView1.Top = 30
  TreeView1.Width = 4755 * sizeAdjust_w
  TreeView1.Height = 3750 * sizeAdjust_h
  
  lsvFileList.Left = TreeView1.Width + 60
  lsvFileList.Top = 30
  lsvFileList.Width = 6345 * sizeAdjust_w
  lsvFileList.Height = 3750 * sizeAdjust_h
  
  'do columns headers
  newHeaderSize = (lsvFileList.Width - 6345) + 4500
  
  If newHeaderSize < 1250 Then
    lsvFileList.ColumnHeaders(1).Width = 2 * newHeaderSize
  Else
    lsvFileList.ColumnHeaders(1).Width = newHeaderSize
  End If
  
  lblDirStats.Left = 30 + 15
  lblDirStats.Top = TreeView1.Height + 60
  lblDirStats.Width = TreeView1.Width - 30
  lblDirStats.Height = 315
  
  lblStats.Left = lsvFileList.Left + 15
  lblStats.Top = lblDirStats.Top
  lblStats.Width = lsvFileList.Width - 30
  lblStats.Height = 315
  
  lblFilePathName.Left = 30 + 15
  lblFilePathName.Top = TreeView1.Height + lblStats.Height + 105
  lblFilePathName.Width = lblStats.Width + lblDirStats.Width + 60
  lblFilePathName.Height = 315
  
  cmdRefresh.Left = lblFilePathName.Left + lblFilePathName.Width - 255
  cmdRefresh.Top = lblFilePathName.Top + 30
End Sub

Private Sub UserControl_Terminate()
  tmrVerify.Enabled = False
End Sub

Public Sub Refresh(Optional ByVal currentFileOnly As Boolean = False, Optional ByVal currentFileSizeChange As Long = 0)
  If currentFileOnly Then
    Dim loop1 As Long, t_size As Long, ts_size As String
    
    For loop1 = 1 To lsvFileList.ListItems.Count
      With lsvFileList.ListItems.Item(loop1)
        If .Text = FileName Then
          t_size = .SubItems(1) + currentFileSizeChange
          ts_size = Format(t_size, "#,###,###,##0")
          
          .SubItems(1) = ts_size
          .SubItems(2) = Chr$(64 + Len(ts_size)) & ts_size
          
          m_Size = m_Size + currentFileSizeChange
          
          lblStats = d_Count & " sub-Directories" & " / " & m_Count & " Files @ " & Format(m_Size, "##,###,###,##0 bytes")
          
          Exit Sub
        End If
      End With
    Next loop1
  Else
    Refresh_DirectoryTree
  End If
End Sub

