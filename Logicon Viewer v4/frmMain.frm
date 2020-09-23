VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logicon File Viewer v4.f (freeware) - Warning! If you have to ask what this program does don't use it!"
   ClientHeight    =   10650
   ClientLeft      =   195
   ClientTop       =   225
   ClientWidth     =   15360
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10650
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdInsert 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14820
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Insert block into file."
      Top             =   1740
      Width           =   405
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14340
      Picture         =   "frmMain.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Delete block from file."
      Top             =   1740
      Width           =   405
   End
   Begin VB.CommandButton cmdBackup 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14820
      Picture         =   "frmMain.frx":0D6E
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Make backup copy of file."
      Top             =   2820
      Width           =   405
   End
   Begin VB.CommandButton cmdViewDetails 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14340
      Picture         =   "frmMain.frx":12A0
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "View file details."
      Top             =   1200
      Width           =   405
   End
   Begin VB.Timer tmrCursorBlink 
      Interval        =   250
      Left            =   300
      Top             =   5760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12840
      TabIndex        =   25
      Top             =   1110
      Width           =   1410
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "QuickCalc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   28
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   11
         Left            =   60
         TabIndex        =   27
         ToolTipText     =   "Activates QuickCalc and shows hexadecimal result."
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   10
         Left            =   60
         TabIndex        =   26
         ToolTipText     =   "Activates QuickCalc and shows decimal value result."
         Top             =   420
         Width           =   1290
      End
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   90
      Top             =   5070
   End
   Begin MSFlexGridLib.MSFlexGrid flxOffset 
      CausesValidation=   0   'False
      Height          =   6075
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   10716
      _Version        =   393216
      Rows            =   27
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorFixed  =   -2147483632
      ForeColorFixed  =   0
      BackColorSel    =   8421504
      BackColorBkg    =   0
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
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
   Begin VB.CommandButton cmdSetFilter 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14820
      Picture         =   "frmMain.frx":17D2
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Set file filter."
      Top             =   2280
      Width           =   405
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   12840
      TabIndex        =   6
      Top             =   2130
      Width           =   1410
      Begin VB.CheckBox chkArchive 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archive"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Indicates if file is marked Archive."
         Top             =   180
         Width           =   1290
      End
      Begin VB.CheckBox chkRead 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Read-Only"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Indicates if file is marked Read-Only"
         Top             =   1170
         Width           =   1290
      End
      Begin VB.CheckBox chkHidden 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hidden"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Indicates if file is marked Hidden."
         Top             =   510
         Width           =   1290
      End
      Begin VB.CheckBox chkSystem 
         BackColor       =   &H00C0C0C0&
         Caption         =   "System"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Indicates if file is marked System."
         Top             =   840
         Width           =   1290
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6075
      LargeChange     =   2
      Left            =   15060
      Max             =   1000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4500
      Width           =   165
   End
   Begin VB.CommandButton cmdScrapNow 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14340
      Picture         =   "frmMain.frx":1A84
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Scrap any changes made."
      Top             =   3300
      Width           =   405
   End
   Begin VB.CheckBox chkOLock 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   14340
      Picture         =   "frmMain.frx":1B86
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Lock offset at which newly selected file first displays. (Used for comparing files)."
      Top             =   2280
      Width           =   405
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   14820
      Picture         =   "frmMain.frx":1C88
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Initiate a search (press F3 to continue a search)."
      Top             =   1200
      Width           =   405
   End
   Begin VB.CommandButton cmdSaveNow 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14820
      Picture         =   "frmMain.frx":1D8A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Save any changes made."
      Top             =   3300
      Width           =   405
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   11355
      TabIndex        =   11
      Top             =   1110
      Width           =   1425
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   4
         Left            =   60
         TabIndex        =   29
         Top             =   2220
         Width           =   1290
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   9
         Left            =   60
         TabIndex        =   21
         ToolTipText     =   "Cursor Address"
         Top             =   180
         Width           =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF80&
         BorderWidth     =   5
         Index           =   0
         X1              =   315
         X2              =   90
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   8
         Left            =   60
         TabIndex        =   19
         Top             =   570
         Width           =   300
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   7
         Left            =   390
         TabIndex        =   18
         Top             =   570
         Width           =   300
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   6
         Left            =   720
         TabIndex        =   17
         Top             =   570
         Width           =   300
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   5
         Left            =   1050
         TabIndex        =   16
         Top             =   570
         Width           =   300
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   60
         TabIndex        =   15
         Top             =   1890
         Width           =   1290
      End
      Begin VB.Label lblValue 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   735
         TabIndex        =   13
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label lblValue 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   1230
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   5
         Index           =   1
         X1              =   660
         X2              =   90
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   5
         Index           =   5
         X1              =   660
         X2              =   90
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   5
         Index           =   4
         X1              =   1320
         X2              =   90
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   5
         Index           =   2
         X1              =   1320
         X2              =   90
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   5
         Index           =   3
         X1              =   1320
         X2              =   90
         Y1              =   1110
         Y2              =   1110
      End
   End
   Begin Project1.Explorer Explorer1 
      Height          =   4470
      Left            =   90
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   8361
   End
   Begin MSFlexGridLib.MSFlexGrid flxHex 
      CausesValidation=   0   'False
      Height          =   6075
      Left            =   1320
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4500
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   10716
      _Version        =   393216
      Rows            =   27
      Cols            =   33
      FixedCols       =   0
      BackColor       =   0
      ForeColor       =   16776960
      ForeColorFixed  =   -2147483640
      BackColorSel    =   8421504
      BackColorBkg    =   0
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
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
   Begin MSFlexGridLib.MSFlexGrid flxASC 
      CausesValidation=   0   'False
      Height          =   6075
      Left            =   10320
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4500
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   10716
      _Version        =   393216
      Rows            =   27
      Cols            =   17
      FixedCols       =   0
      BackColor       =   0
      ForeColor       =   10526720
      ForeColorFixed  =   -2147483640
      BackColorSel    =   8421504
      BackColorBkg    =   0
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
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
   Begin VB.Label lblFileStats 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   11370
      TabIndex        =   22
      ToolTipText     =   "Current active file type."
      Top             =   4110
      Width           =   3855
   End
   Begin VB.Image lblLocked 
      Height          =   480
      Left            =   11520
      Picture         =   "frmMain.frx":1E8C
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   12420
      Picture         =   "frmMain.frx":22CE
      Stretch         =   -1  'True
      ToolTipText     =   "About this program."
      Top             =   60
      Width           =   2790
   End
   Begin VB.Image imgPicture 
      Height          =   3660
      Left            =   5370
      Picture         =   "frmMain.frx":3014
      Top             =   4680
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.Label lblFileStats 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   11370
      TabIndex        =   0
      ToolTipText     =   "Current active file size."
      Top             =   3750
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const mylabel As String = "eftjhofe!cz!Ujn!Ibsqvs"
Private oldCursorRow As Long, oldCursorCol As Long

Private Sub chkArchive_Click()
  If chkArchive.Value = 1 Then
    If noSound = 0 Then DXSound.Play_Sound 2
  Else
    If noSound = 0 Then DXSound.Play_Sound 3
  End If
End Sub

Private Sub chkHidden_Click()
  If chkHidden.Value = 1 Then
    If noSound = 0 Then DXSound.Play_Sound 2
  Else
    If noSound = 0 Then DXSound.Play_Sound 3
  End If
End Sub

Private Sub chkOLock_Click()
  If chkOLock.Value = 1 Then
    If noSound = 0 Then DXSound.Play_Sound 2
  Else
    If noSound = 0 Then DXSound.Play_Sound 3
  End If
End Sub

Private Sub chkRead_Click()
  If chkRead.Value = 1 Then
    If noSound = 0 Then DXSound.Play_Sound 2
  Else
    If noSound = 0 Then DXSound.Play_Sound 3
  End If
End Sub

Private Sub chkSystem_Click()
  If chkSystem.Value = 1 Then
    If noSound = 0 Then DXSound.Play_Sound 2
  Else
    If noSound = 0 Then DXSound.Play_Sound 3
  End If
End Sub

Private Sub cmdBackup_Click()
  On Error GoTo gotErr
  
  If noSound = 0 Then DXSound.Play_Sound 1
  
  If vFileSize <= 0 Then Exit Sub
  
  SaveBlock
  
  Close #1
  
  FileCopy vFilePath & vFileName, vFilePath & vFileName & ".bak"
  
  Open vFilePath & vFileName For Binary As #1
  
  Explorer1.Refresh
  
  Exit Sub
  
gotErr:
  MsgBox "Critical error while copying file!", vbOKOnly, "File Error"
End Sub

Private Sub cmdDelete_Click()
  Form_KeyDown vbKeyDelete, 0
End Sub

Private Sub cmdFind_Click(Index As Integer)
  Dim SearchOffset As Long, SearchLen As Long, SearchStop As Long, tempstring As String
  Dim HoldOffset As Long, SearchMatch As Boolean, ReplaceCount As Long
  Dim SByte1 As Long, SByte2 As Long, SByte3 As Long, SByte4 As Long
  Dim loop1 As Long, temp1 As Long, temp2 As Byte, temp3 As Byte, tempstring2 As String
  
  On Error Resume Next
  
  If noSound = 0 Then DXSound.Play_Sound 2
  
  If vFileSize <= 0 Then Exit Sub
  
  If SearchString <> "" And Index = -9 Then
    SearchStart = 2
  Else
    frmSearch.Show vbModal
    
    If noSound = 0 Then DXSound.Play_Sound 3
  End If
  
  If SearchString = "" Or SearchStart = 0 Then Exit Sub
  
  Me.Enabled = False
  
  SaveBlock
  
  HoldOffset = BaseOffset
  
  If SearchStart <> 2 Then
    SearchOffset = 0
    
    If BaseOffset <> 0 Then
      BaseOffset = 0
      
      LoadBlock
    End If
  Else
    SearchOffset = BaseOffset + PageOffset + LineOffset * prjLineWidth + RowOffset + 1
  End If
  
  lblFileStats(1).Caption = Format(SearchOffset, "#,###,###,##0") & " Bytes"
  
  SearchOK = True
  SearchMatch = False
  ReplaceCount = 0
  
  On Error GoTo valuemismatch
  
  Select Case SearchMode
    Case 0
      SearchString = LCase(SearchString)
      tempstring = SearchString
      tempstring2 = ReplaceString
      
      SearchLen = Len(SearchString)
      
      SByte1 = Asc(Left$(SearchString, 1))
      
      If SByte1 >= 97 And SByte1 <= 122 Then
        SByte2 = SByte1 - 32
      Else
        SByte2 = SByte1
      End If
    Case 1
      SearchString = LCase(SearchString)
      SearchLen = 2 * Len(SearchString)
      tempstring = ""
      tempstring2 = ""
      
      If SearchStart = 3 Then
        For loop1 = 1 To SearchLen \ 2
          tempstring = tempstring & Mid$(SearchString, loop1, 1) & Chr$(0)
          tempstring2 = tempstring2 & Mid$(ReplaceString, loop1, 1) & Chr$(0)
        Next loop1
      Else
        For loop1 = 1 To SearchLen \ 2
          tempstring = tempstring & Mid$(SearchString, loop1, 1) & Chr$(0)
        Next loop1
      End If
      
      SByte1 = Asc(Left$(tempstring, 1))
      
      If SByte1 >= 97 And SByte1 <= 122 Then
        SByte2 = SByte1 - 32
      Else
        SByte2 = SByte1
      End If
    Case 2
      SearchLen = 1
      
      SByte1 = val(SearchString)
      
      If SByte1 < -128 Or SByte1 > 255 Then Error 1
      If SByte1 < 0 Then SByte1 = 256 + SByte1
    Case 3
      SearchLen = 2
      
      SByte1 = val(SearchString)
      
      If SByte1 < -32768 Or SByte1 > 65535 Then Error 1
      If SByte1 < 0 Then SByte1 = 65536 + SByte1
      
      SByte2 = SByte1 \ 256
      SByte1 = SByte1 Mod 256
    Case 4
      SearchLen = 4
      
      SByte1 = val(SearchString)
      
      If SByte1 < 0 Then
        SByte1 = (-SByte1) - 1
        
        SByte4 = SByte1 \ 16777216
        SByte1 = SByte1 Mod 16777216
        
        SByte3 = SByte1 \ 65536
        SByte1 = SByte1 Mod 65536
        
        SByte2 = SByte1 \ 256
        SByte1 = SByte1 Mod 256
        
        SByte4 = 256 - (SByte4 + 1)
        SByte3 = 256 - (SByte3 + 1)
        SByte2 = 256 - (SByte2 + 1)
        SByte1 = 256 - (SByte1 + 1)
      Else
        SByte4 = SByte1 \ 16777216
        SByte1 = SByte1 Mod 16777216
        
        SByte3 = SByte1 \ 65536
        SByte1 = SByte1 Mod 65536
        
        SByte2 = SByte1 \ 256
        SByte1 = SByte1 Mod 256
      End If
  End Select
  
  Do While SearchOK
    If (BaseOffset + BlockSize - SearchLen) <= SearchOffset Then
      If SearchStart = 3 And SearchMatch Then
        Put #1, BaseOffset + 1, FileData
        
        SearchMatch = False
      End If
      
      BaseOffset = (SearchOffset \ prjLineWidth) * prjLineWidth
      LoadBlock
    End If
    
    SearchStop = BaseOffset + BlockSize - SearchLen
    
    Select Case SearchMode
      Case 0, 1
        Do While SearchOffset <= SearchStop
          temp2 = FileData(SearchOffset - BaseOffset)
          
          If SByte1 = temp2 Or SByte2 = temp2 Then
            loop1 = SearchLen - 1
            
            Do While loop1 > 0
              temp2 = FileData(SearchOffset - BaseOffset + loop1)
              SByte3 = Asc(Mid$(tempstring, loop1 + 1, 1))
              
              If SByte3 >= 97 And SByte3 <= 122 Then
                SByte4 = SByte3 - 32
              Else
                SByte4 = SByte3
              End If
              
              If temp2 <> SByte3 And temp2 <> SByte4 Then Exit Do
              
              loop1 = loop1 - 1
            Loop
            
            If loop1 = 0 Then
              If SearchStart = 3 Then
                ReplaceCount = ReplaceCount + 1
                
                For loop1 = 0 To SearchLen - 1
                  FileData(SearchOffset - BaseOffset + loop1) = Asc(Mid$(tempstring2, loop1 + 1, 1))
                  
                  SearchMatch = True
                Next loop1
              Else
                SearchOK = False
                SearchMatch = True
                
                Exit Do
              End If
            End If
          End If
          
          SearchOffset = SearchOffset + 1
        Loop
      Case 2
        Do While SearchOffset <= SearchStop
          If SByte1 = FileData(SearchOffset - BaseOffset) Then
            SearchOK = False
            SearchMatch = True
            
            Exit Do
          End If
          
          SearchOffset = SearchOffset + 1
        Loop
      Case 3
        Do While SearchOffset <= SearchStop
          If SByte1 = FileData(SearchOffset - BaseOffset) Then
            If SByte2 = FileData(SearchOffset - BaseOffset + 1) Then
              SearchOK = False
              SearchMatch = True
              
              Exit Do
            End If
          End If
          
          SearchOffset = SearchOffset + 1
        Loop
      Case 4
        Do While SearchOffset <= SearchStop
          If SByte1 = FileData(SearchOffset - BaseOffset) Then
            If SByte2 = FileData(SearchOffset - BaseOffset + 1) Then
              If SByte3 = FileData(SearchOffset - BaseOffset + 2) Then
                If SByte4 = FileData(SearchOffset - BaseOffset + 3) Then
                  SearchOK = False
                  SearchMatch = True
                  
                  Exit Do
                End If
              End If
            End If
          End If
          
          SearchOffset = SearchOffset + 1
        Loop
    End Select
    
    If SearchOK Then
      If SearchOffset > (vFileSize - SearchLen) Then SearchOK = False
      
      lblFileStats(1).Caption = Format(SearchOffset, "#,###,###,##0")
      
      DoEvents
    End If
  Loop
  
  lblFileStats(1).Caption = Format(vFileSize, "#,###,###,##0") & " Bytes"
  
  Me.Enabled = True
  
  If SearchStart = 3 Then
    If SearchMatch Then Put #1, BaseOffset + 1, FileData
      
    BaseOffset = HoldOffset
    
    LoadBlock
    UpdateDisplay
    
    MsgBox ReplaceCount & " count(s) of (" & SearchString & ") were replaced with (" & ReplaceString & ")", vbOKOnly
  ElseIf SearchMatch Then
    If BaseOffset = HoldOffset Then
      If (SearchOffset < BaseOffset + PageOffset) Or (SearchOffset + SearchLen > BaseOffset + PageOffset + prjBlockSize) Then
        PageOffset = (((SearchOffset - BaseOffset) \ prjLineWidth) * prjLineWidth)
      End If
    Else
      PageOffset = (((SearchOffset - BaseOffset) \ prjLineWidth) * prjLineWidth)
    End If
    
    RowOffset = SearchOffset - (BaseOffset + PageOffset)
    LineOffset = RowOffset \ prjLineWidth
    RowOffset = RowOffset Mod prjLineWidth
    
    If SearchMode < 2 Then
      eMode = True
    Else
      eMode = False
    End If
      
    UpdateDisplay
    
    temp1 = SearchLen \ prjLineWidth
    SearchLen = SearchLen Mod prjLineWidth
  Else
    BaseOffset = HoldOffset
    
    LoadBlock
    
    MsgBox "No further matches found for :" & Chr$(13) & Chr$(13) & SearchString, vbOKOnly
  End If
  
  ShowCursor
  
  Exit Sub
  
valuemismatch:
  MsgBox "The value is invalid and does not match the type selected!", vbOKOnly, "ERROR!"
  
  BaseOffset = HoldOffset
    
  LoadBlock
  
  Me.Enabled = True
End Sub

Private Sub cmdInsert_Click()
  Form_KeyDown vbKeyInsert, 0
End Sub

Private Sub cmdSaveNow_Click()
  If noSound = 0 Then DXSound.Play_Sound 1
  
  If vFileName <> "" And FileChange Then SaveBlock
End Sub

Private Sub cmdScrapNow_Click()
  If noSound = 0 Then DXSound.Play_Sound 2
  
  If vFileName <> "" Then
    If MsgBox("Scrap changes to file?", vbYesNo, "Scrapping Changes") = vbYes Then
      If noSound = 0 Then DXSound.Play_Sound 3
      
      LoadBlock
      
      UpdateDisplay
    Else
      If noSound = 0 Then DXSound.Play_Sound 3
    End If
  End If
End Sub

Private Sub cmdSetFilter_Click()
  If noSound = 0 Then DXSound.Play_Sound 2
  
  frmFilter.Show vbModal
  
  If noSound = 0 Then DXSound.Play_Sound 3
End Sub

Private Sub cmdViewDetails_Click()
  frmView.Show vbModal
End Sub

Private Sub Explorer1_ClickSound()
  If noSound = 0 Then DXSound.Play_Sound 2
  If noSound = 1 Then noSound = 0
End Sub

Private Sub Explorer1_FileSelected()
  Dim TempAttrib As Long, goterror As Boolean
  
  noSound = 9
  
  If vFileName <> "" Then
    If FileChange Then SaveBlock
    
    On Error Resume Next
    
    Close #1
    
    If chkArchive.Value = 1 Then TempAttrib = TempAttrib + vbArchive
    If chkRead.Value = 1 Then TempAttrib = TempAttrib + vbReadOnly
    If chkSystem.Value = 1 Then TempAttrib = TempAttrib + vbSystem
    If chkHidden.Value = 1 Then TempAttrib = TempAttrib + vbHidden
    
    If Not FileLock Then SetAttr vFilePath & vFileName, TempAttrib
  End If
  
  ShowCursor True
  
  goterror = False
  
  lblLocked.Visible = False
  
  chkArchive.Enabled = False
  chkRead.Enabled = False
  chkSystem.Enabled = False
  chkHidden.Enabled = False
  chkArchive.Value = 0
  chkRead.Value = 0
  chkSystem.Value = 0
  chkHidden.Value = 0
  
  cmdSaveNow.Enabled = False
  cmdScrapNow.Enabled = False
  
  LastVScroll = 0
  BlockSize = 0
  vFileSize = 0
  LineOffset = 0
  RowOffset = 0
  RowSubOffset = False
  FileLock = False
  
  VScroll1.Value = 0
  
  On Error GoTo badfile
  vFileName = Explorer1.FileName
  vFilePath = Explorer1.FilePath
  If Right$(vFilePath, 1) <> "\" Then vFilePath = vFilePath & "\"
  vFileSize = Explorer1.FileSize
  TempAttrib = Explorer1.FileAttributes
  
  lblFileStats(0).Caption = Explorer1.FileType
  lblFileStats(1).Caption = Format(vFileSize, "##,###,###,##0") & " Bytes"
  
  If TempAttrib And vbArchive Then chkArchive.Value = 1
  If TempAttrib And vbReadOnly Then chkRead.Value = 1
  If TempAttrib And vbSystem Then chkSystem.Value = 1
  If TempAttrib And vbHidden Then chkHidden.Value = 1
  
  On Error GoTo lockedfile
  
  SetAttr vFilePath & vFileName, vbNormal
  
  Open vFilePath & vFileName For Binary As #1
  
  If Not goterror Then
    chkArchive.Enabled = True
    chkRead.Enabled = True
    chkSystem.Enabled = True
    chkHidden.Enabled = True
  End If
  
  If chkOLock.Value = 0 Or (BaseOffset + PageOffset >= vFileSize) Then
    BaseOffset = 0
    PageOffset = 0
  End If
  
  LoadBlock
  UpdateDisplay
  ShowCursor
  
  noSound = 0
  
  Exit Sub

lockedfile:
  If goterror Then
badfile:
    MsgBox "Check drive..." & Chr$(13) & Chr$(13) & "Could not access the selected file!", vbOKOnly, "ERROR!"
    
    vFileName = ""
    
    FileChange = False
    cmdScrapNow.Enabled = False
    cmdSaveNow.Enabled = False
      
    cmdViewDetails.Enabled = False
    cmdFind(0).Enabled = False
    cmdDelete.Enabled = False
    cmdInsert.Enabled = False
  
    UpdateDisplay
    
    noSound = 0
  Else
    lblLocked.Visible = True
    lblLocked.ToolTipText = "Write Locked File"
    
    FileLock = True
    
    cmdScrapNow.Enabled = False
    cmdSaveNow.Enabled = False
      
    cmdDelete.Enabled = False
    cmdInsert.Enabled = False
    
    goterror = True
    
    Resume Next
  End If
End Sub

Private Sub Explorer1_FileRenamed(oldFilePathName As String, newFilePath As String, newFileName As String)
  If vFileName <> "" Then
    vFilePath = newFilePath
    If Right$(vFilePath, 1) <> "\" Then vFilePath = vFilePath & "\"
    
    vFileName = newFileName
    
    Open vFilePath & vFileName For Binary As #1
  End If
End Sub

Private Sub Explorer1_FileEditStart(oldPathName As String)
  'release file handle before continuing with the rename
  
  If vFileName <> "" Then
    On Error Resume Next
    
    Close #1
  End If
End Sub

Private Sub Explorer1_PathRenamed(oldPath As String, newPath As String)
  If oldPath & "\" = vFilePath Then vFilePath = newPath & "\"
End Sub

Private Sub flxASC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If flxASC.Row > 0 And flxASC.Row <= prjLineHeight Then
    If flxASC.Col < prjLineWidth Then
      LineOffset = flxASC.Row - 1
      RowOffset = flxASC.Col
      RowSubOffset = False
      
      If eMode = False Then
        eMode = True
        
        flxHex.ForeColor = &HA0A000
        flxASC.ForeColor = &HFFFF00
      End If
      
      ShowCursor
    End If
  End If
End Sub

Private Sub flxHex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If flxHex.Row > 0 And flxHex.Row <= prjLineHeight Then
    If flxHex.Col < prjLineWidth * 2 Then
      LineOffset = flxHex.Row - 1
      RowOffset = flxHex.Col
      
      If RowOffset Mod 2 = 0 Then
        RowSubOffset = False
      Else
        RowSubOffset = True
      End If
      
      RowOffset = RowOffset \ 2
      
      If eMode Then
        eMode = False
        
        flxHex.ForeColor = &HFFFF00
        flxASC.ForeColor = &HA0A000
      End If
      
      ShowCursor
    End If
  End If
End Sub

Private Sub flxOffset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim loop1 As Long
  
  aMode = Not aMode
  
  With frmMain.flxHex
    If aMode Then
      For loop1 = 0 To prjLineWidth - 1
        If loop1 < 10 Then
          .TextMatrix(0, loop1 * 2) = ""
          .TextMatrix(0, loop1 * 2 + 1) = loop1
        Else
          .TextMatrix(0, loop1 * 2) = 1
          .TextMatrix(0, loop1 * 2 + 1) = loop1 - 10
        End If
      Next loop1
    Else
      For loop1 = 0 To prjLineWidth - 1
        .TextMatrix(0, loop1 * 2) = MakeHexShortH(loop1)
        .TextMatrix(0, loop1 * 2 + 1) = MakeHexShortL(loop1)
      Next loop1
    End If
  End With
  
  UpdateDisplay
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim curPos As Long
  
  If vFileName = "" Or vFileSize <= 0 Then
    KeyCode = 0
    Shift = 0
    
    Exit Sub
  End If
  
  Select Case KeyCode
    Case vbKeyInsert
      SaveBlock
      
      frmInsert.Show vbModal
      
      LoadBlock
      
      UpdateDisplay
      
      KeyCode = 0
      Shift = 0
    Case vbKeyDelete
      SaveBlock
      
      frmDelete.Show vbModal
      
      LoadBlock
      
      UpdateDisplay
      
      KeyCode = 0
      Shift = 0
    Case vbKeyTab
      
      KeyCode = 0
      Shift = 0
      
      eMode = Not eMode
      
      If eMode Then
        flxHex.ForeColor = &HA0A000
        flxASC.ForeColor = &HFFFF00
      Else
        flxHex.ForeColor = &HFFFF00
        flxASC.ForeColor = &HA0A000
      End If
      
      RowSubOffset = False
      
      ShowCursor
    Case vbKeyF3
      
      KeyCode = 0
      Shift = 0
      
      cmdFind_Click (-9)
    Case vbKeyUp
            
      If LineOffset > 0 Then
        LineOffset = LineOffset - 1
      ElseIf PageOffset > 0 Then
        PageOffset = PageOffset - prjLineWidth
        
        UpdateDisplay
      ElseIf BaseOffset > 0 Then
        SaveBlock
        
        PageOffset = PageOffset + BaseOffset
        
        BaseOffset = BaseOffset - 24000
        
        If BaseOffset < 0 Then BaseOffset = 0
        
        PageOffset = PageOffset - BaseOffset - prjLineWidth
        If PageOffset < 0 Then PageOffset = 0
        
        LoadBlock
      End If
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyDown
            
      If LineOffset < prjLineHeight - 1 Then
        LineOffset = LineOffset + 1
      ElseIf PageOffset < BlockSize - prjBlockSize Then
        PageOffset = PageOffset + prjLineWidth
        
        UpdateDisplay
      ElseIf BaseOffset + PageOffset + prjLineWidth < vFileSize And BlockSize = 48000 Then
        SaveBlock
        
        BaseOffset = BaseOffset + 24000
        PageOffset = PageOffset - 24000 + prjLineWidth
        
        LoadBlock
      End If
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyLeft
      
      If Shift And vbCtrlMask Then
        curPos = BaseOffset + PageOffset + (LineOffset * prjLineWidth) + RowOffset
        
        If RowOffset > 0 Then
          RowOffset = RowOffset - 1
        ElseIf BaseOffset + PageOffset + LineOffset > 0 Then
          RowOffset = (prjLineWidth - 1)
          
          Form_KeyDown vbKeyUp, 0
        End If
        
        If curPos <> BaseOffset + PageOffset + (LineOffset * prjLineWidth) + RowOffset Then
          curPos = PageOffset + (LineOffset * prjLineWidth) + RowOffset
          
          FileData(curPos + 1) = FileData(curPos)
          
          FileChange = True
          cmdScrapNow.Enabled = True
          cmdSaveNow.Enabled = True
          
          UpdateDisplay
        End If
      ElseIf eMode Then
        If RowOffset > 0 Then
          RowOffset = RowOffset - 1
        ElseIf BaseOffset + PageOffset + LineOffset > 0 Then
          RowOffset = (prjLineWidth - 1)
          
          Form_KeyDown vbKeyUp, 0
        End If
      Else
        RowSubOffset = Not RowSubOffset
        
        If RowSubOffset Then
          If RowOffset > 0 Then
            RowOffset = RowOffset - 1
          ElseIf BaseOffset + PageOffset + LineOffset > 0 Then
            RowOffset = (prjLineWidth - 1)
            
            Form_KeyDown vbKeyUp, 0
          End If
        End If
      End If
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyRight
       
      If Shift And vbCtrlMask Then
        curPos = BaseOffset + PageOffset + (LineOffset * prjLineWidth) + RowOffset
        
        If RowOffset < (prjLineWidth - 1) Then
          RowOffset = RowOffset + 1
        ElseIf BaseOffset + PageOffset + (LineOffset * prjLineWidth) < vFileSize Then
          RowOffset = 0
          
          Form_KeyDown vbKeyDown, 0
        End If
        
        If curPos <> BaseOffset + PageOffset + (LineOffset * prjLineWidth) + RowOffset Then
          curPos = PageOffset + (LineOffset * prjLineWidth) + RowOffset
          
          FileData(curPos - 1) = FileData(curPos)
          
          FileChange = True
          cmdScrapNow.Enabled = True
          cmdSaveNow.Enabled = True
          
          UpdateDisplay
        End If
      ElseIf eMode Then
        If RowOffset < (prjLineWidth - 1) Then
          RowOffset = RowOffset + 1
        ElseIf BaseOffset + PageOffset + (LineOffset * prjLineWidth) < vFileSize Then
          RowOffset = 0
          
          Form_KeyDown vbKeyDown, 0
        End If
      Else
        RowSubOffset = Not RowSubOffset
        
        If Not RowSubOffset Then
          If RowOffset < (prjLineWidth - 1) Then
            RowOffset = RowOffset + 1
          ElseIf BaseOffset + PageOffset + (LineOffset * prjLineWidth) < vFileSize Then
            RowOffset = 0
            
            Form_KeyDown vbKeyDown, 0
          End If
        End If
      End If
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyHome
            
      LineOffset = 0
      RowOffset = 0
      RowSubOffset = False
      PageOffset = 0
      
      If BaseOffset > 0 Then
        SaveBlock
        
        BaseOffset = 0
        
        LoadBlock
      End If
      
      UpdateDisplay
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyPageUp
            
      If PageOffset > 0 Or BaseOffset > 0 Then
        PageOffset = PageOffset - prjBlockSize
        
        If PageOffset < 0 Then
          If BaseOffset > 0 Then
            PageOffset = BaseOffset + PageOffset
            
            SaveBlock
            
            BaseOffset = BaseOffset - 24000
            
            If BaseOffset < 0 Then BaseOffset = 0
            
            LoadBlock
            
            PageOffset = PageOffset - BaseOffset
            If PageOffset < 0 Then PageOffset = 0
          Else
            PageOffset = 0
          End If
        End If
        
        UpdateDisplay
      ElseIf LineOffset > 0 Then
        LineOffset = 0
      End If
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyPageDown
            
      If BaseOffset + PageOffset + prjBlockSize < vFileSize Then
        PageOffset = PageOffset + prjBlockSize
        
        If PageOffset + prjBlockSize >= BlockSize Then
          If BaseOffset + BlockSize < vFileSize Then
            PageOffset = PageOffset + BaseOffset
            
            SaveBlock
            
            BaseOffset = BaseOffset + 24000
            
            If vFileSize - BaseOffset < 48000 Then BaseOffset = (vFileSize \ prjLineWidth) * prjLineWidth - 24000
            If BaseOffset < 0 Then BaseOffset = 0
            
            LoadBlock
            
            PageOffset = PageOffset - BaseOffset
          End If
          
          If PageOffset < 0 Then
            PageOffset = 0
          Else
            If PageOffset + prjBlockSize >= BlockSize Then PageOffset = ((BlockSize + (prjLineWidth - 1)) \ prjLineWidth) * prjLineWidth - prjBlockSize
            If PageOffset < 0 Then PageOffset = 0
          End If
        End If
        
        UpdateDisplay
      ElseIf LineOffset < prjLineHeight - 1 Then
        LineOffset = prjLineHeight - 1
      End If
      
      KeyCode = 0
      Shift = 0
      ShowCursor
    Case vbKeyEnd
            
      If BaseOffset + BlockSize < vFileSize Then
        SaveBlock
        
        BaseOffset = ((vFileSize \ prjLineWidth) * prjLineWidth) - 24000
        If BaseOffset < 0 Then BaseOffset = 0
        
        LoadBlock
      End If
      
      PageOffset = ((BlockSize + (prjLineWidth - 1)) \ prjLineWidth) * prjLineWidth - prjBlockSize
      If PageOffset < 0 Then PageOffset = 0
      
      RowOffset = vFileSize - BaseOffset - PageOffset - 1
      If RowOffset < 0 Then RowOffset = 0
      
      LineOffset = RowOffset \ prjLineWidth
      RowOffset = RowOffset Mod prjLineWidth
      RowSubOffset = True
      
      UpdateDisplay
      
      KeyCode = 0
      Shift = 0
      ShowCursor
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim valh As Long, vall As Long, Currentpos As Long
  
  If vFileName = "" Or vFileSize <= 0 Or FileLock Then
    KeyAscii = 0
    
    Exit Sub
  End If
  
  Currentpos = PageOffset + (LineOffset * prjLineWidth) + RowOffset
  
  If eMode Then
    If KeyAscii < 32 Then
      KeyAscii = 0
      
      Exit Sub
    End If
  Else
    valh = FileData(Currentpos)
    
    If RowSubOffset Then
      valh = valh \ 16
      
      If KeyAscii >= 48 And KeyAscii <= 57 Then
        vall = KeyAscii - 48
      ElseIf KeyAscii >= 65 And KeyAscii <= 70 Then
        vall = KeyAscii - 55
      ElseIf KeyAscii >= 97 And KeyAscii <= 102 Then
        vall = KeyAscii - 87
      Else
        KeyAscii = 0
        
        Exit Sub
      End If
    Else
      vall = valh Mod 16
      
      If KeyAscii >= 48 And KeyAscii <= 57 Then
        valh = KeyAscii - 48
      ElseIf KeyAscii >= 65 And KeyAscii <= 70 Then
        valh = KeyAscii - 55
      ElseIf KeyAscii >= 97 And KeyAscii <= 102 Then
        valh = KeyAscii - 87
      Else
        KeyAscii = 0
        
        Exit Sub
      End If
    End If
    
    KeyAscii = valh * 16 + vall
  End If
  
  FileChange = True
  cmdScrapNow.Enabled = True
  cmdSaveNow.Enabled = True
    
  FileData(Currentpos) = KeyAscii
  
  UpdateDisplay
  
  Form_KeyDown vbKeyRight, 0
  
  KeyAscii = 0
End Sub

Private Sub Form_Load()
  SizeForm Me
  
  SearchString = ""
  SearchMode = 0
  vFileSize = 0
  BlockSize = 0
  vFileName = ""
  
  LockOut = False
  noSound = 1
  
  On Error Resume Next
  
  Explorer1.flvPicture imgPicture.Picture
  Explorer1.FileFilter = "*.*"
  
  DXSound.Init_DXSound Me, 3
  
  DXSound.Load_SoundBuffer 1, "bClick1.wav"
  DXSound.Load_SoundBuffer 2, "bClick2a.wav"
  DXSound.Load_SoundBuffer 3, "bClick2b.wav"
  
  DXSound.Change_SoundSettings 1, 25000
  DXSound.Change_SoundSettings 2, 25000
  DXSound.Change_SoundSettings 3, 25000
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim TempAttrib As Long
  
  If vFileName <> "" Then
    SaveBlock
    
    Close #1
    
    TempAttrib = 0
    
    If chkArchive.Value = 1 Then TempAttrib = TempAttrib + vbArchive
    If chkRead.Value = 1 Then TempAttrib = TempAttrib + vbReadOnly
    If chkSystem.Value = 1 Then TempAttrib = TempAttrib + vbSystem
    If chkHidden.Value = 1 Then TempAttrib = TempAttrib + vbHidden
    
    If Not FileLock Then SetAttr vFilePath & vFileName, TempAttrib
  End If
  
  DXSound.CleanUp_DXSound
End Sub

Private Sub Image1_Click()
  If noSound = 0 Then DXSound.Play_Sound 2
  
  frmSplash.TitleBar = Me.Caption & vbNewLine & vbNewLine & "contact by e-mail at tim.harpur@sympatico.ca"
  frmSplash.Show vbModal
  
  If noSound = 0 Then DXSound.Play_Sound 3
End Sub

Private Sub lblValue_Click(Index As Integer)
  Dim Currentpos As Long
  Dim val1 As Long, val2 As Long, val3 As Long, val4 As Long
  
  If Index = 9 Then
    aMode2 = Not aMode2
    
    GenerateValues
    
    Exit Sub
  ElseIf Index > 4 Then
    If Index >= 10 Then
      If noSound = 0 Then DXSound.Play_Sound 2
      
      frmQuickCalc.Show vbModal
      
      If noSound = 0 Then DXSound.Play_Sound 3
      
      If quickCalcResult < 0 Then
        lblValue(10).Caption = Format(quickCalcResult, "#,###,###,##0")
        
        quickCalcResult = -(quickCalcResult + 1)
          
        val4 = quickCalcResult \ 16777216
        quickCalcResult = quickCalcResult Mod 16777216
        val3 = quickCalcResult \ 65536
        quickCalcResult = quickCalcResult Mod 65536
        val1 = quickCalcResult Mod 256
        val2 = quickCalcResult \ 256
        
        val4 = val4 Xor &HFF
        val3 = val3 Xor &HFF
        val2 = val2 Xor &HFF
        val1 = val1 Xor &HFF
      Else
        lblValue(10).Caption = Format(quickCalcResult, "+#,###,###,##0")
        
        val4 = quickCalcResult \ 16777216
        quickCalcResult = quickCalcResult - val4 * 16777216
        val3 = quickCalcResult \ 65536
        quickCalcResult = quickCalcResult Mod 65536
        val1 = quickCalcResult Mod 256
        val2 = quickCalcResult \ 256
      End If
      
      lblValue(11).Caption = MakeHexShort(val4) & MakeHexShort(val3) & MakeHexShort(val2) & MakeHexShort(val1)
    End If
    
    Exit Sub
  End If
  
  If lblValue(Index) = "--" Then Exit Sub
  
  RangeType = Index
  SourceValue = CLng(lblValue(Index))
  
  If noSound = 0 Then DXSound.Play_Sound 2
  
  frmValue.Show vbModal
  
  If noSound = 0 Then DXSound.Play_Sound 3
  
  If SourceValue <> "Cancel" Then
    Currentpos = PageOffset + (LineOffset * prjLineWidth) + RowOffset
    
    Select Case Index
      Case 0:
        val1 = SourceValue
        
        FileData(Currentpos) = val1
      Case 1:
        If SourceValue < 0 Then SourceValue = 256 + SourceValue
        val1 = SourceValue
        
        FileData(Currentpos) = val1
      Case 2:
        val1 = SourceValue Mod 256
        val2 = SourceValue \ 256
        
        FileData(Currentpos) = val1
        FileData(Currentpos + 1) = val2
      Case 3:
        If SourceValue < 0 Then SourceValue = 65536 + SourceValue
        
        val1 = SourceValue Mod 256
        val2 = SourceValue \ 256
        
        FileData(Currentpos) = val1
        FileData(Currentpos + 1) = val2
      Case 4:
        If SourceValue < 0 Then
          SourceValue = -(SourceValue + 1)
            
          val4 = SourceValue \ 16777216
          SourceValue = SourceValue Mod 16777216
          val3 = SourceValue \ 65536
          SourceValue = SourceValue Mod 65536
          val1 = SourceValue Mod 256
          val2 = SourceValue \ 256
          
          val4 = val4 Xor &HFF
          val3 = val3 Xor &HFF
          val2 = val2 Xor &HFF
          val1 = val1 Xor &HFF
        Else
          val4 = SourceValue \ 16777216
          SourceValue = SourceValue - val4 * 16777216
          val3 = SourceValue \ 65536
          SourceValue = SourceValue Mod 65536
          val1 = SourceValue Mod 256
          val2 = SourceValue \ 256
        End If
        
        FileData(Currentpos) = val1
        FileData(Currentpos + 1) = val2
        FileData(Currentpos + 2) = val3
        FileData(Currentpos + 3) = val4
    End Select
    
    FileChange = True
    cmdScrapNow.Enabled = True
    cmdSaveNow.Enabled = True
  
    UpdateDisplay
  End If
End Sub

Private Sub ShowCursor(Optional ByVal clearOnly As Boolean = False, Optional ByVal quickly As Boolean = True)
  Dim cursorpos As Long, holdRow As Long, holdCol As Long
  
  tmrCursorBlink.Enabled = False
  
  If oldeMode Then
    With flxASC
      If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
      If .Col <> oldCursorCol Then .Col = oldCursorCol
      
      If .CellForeColor <> 0 Then .CellForeColor = 0
      If .CellBackColor <> 0 Then .CellBackColor = 0
    End With
    
    With flxHex
      If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
      If .Col <> oldCursorCol * 2 Then .Col = oldCursorCol * 2
      
      If .CellForeColor <> 0 Then .CellForeColor = 0
      If .CellBackColor <> 0 Then .CellBackColor = 0
    End With
  Else
    With flxASC
      If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
      If .Col <> oldCursorCol \ 2 Then .Col = oldCursorCol \ 2
      
      If .CellForeColor <> 0 Then .CellForeColor = 0
      If .CellBackColor <> 0 Then .CellBackColor = 0
    End With
    
    With flxHex
      If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
      If .Col <> oldCursorCol Then .Col = oldCursorCol
      
      If .CellForeColor <> 0 Then .CellForeColor = 0
      If .CellBackColor <> 0 Then .CellBackColor = 0
    End With
  End If
  
  holdRow = oldCursorRow
  holdCol = oldCursorCol
  
  'validate new cursor position
  If BaseOffset + PageOffset + (LineOffset * prjLineWidth) + RowOffset >= vFileSize Then
    cursorpos = vFileSize - BaseOffset - PageOffset - 1
    
    If cursorpos < 0 Then cursorpos = 0
    
    LineOffset = cursorpos \ prjLineWidth
    RowOffset = (cursorpos Mod prjLineWidth)
    RowSubOffset = False
  End If
  
  If LineOffset < 0 Then LineOffset = 0
  If RowOffset < 0 Then RowOffset = 0
  
  oldCursorRow = LineOffset
  
  If eMode Then
    oldCursorCol = RowOffset
  Else
    If RowSubOffset Then
      oldCursorCol = RowOffset * 2 + 1
    Else
      oldCursorCol = RowOffset * 2
    End If
  End If
  
  oldeMode = eMode
  
  If holdRow <> oldCursorRow Then
    With flxOffset
      If .Row <> holdRow + 1 Then .Row = holdRow + 1
      If .CellForeColor <> 0 Then .CellForeColor = 0
      If .CellBackColor <> 0 Then .CellBackColor = 0
    End With
  End If
  
  If clearOnly = False And vFileSize > 0 Then
    If eMode Then
      With flxHex
        If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
        If .Col <> oldCursorCol * 2 Then .Col = oldCursorCol * 2
        
        If .CellBackColor <> &H303030 Then
          .CellBackColor = &H303030
          .CellForeColor = &HC0C0&
        End If
      End With
      
      With flxASC
        If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
        If .Col <> oldCursorCol Then .Col = oldCursorCol
        
        If .CellBackColor <> &H505050 Then
          .CellBackColor = &H505050
          .CellForeColor = &HFFFF&
          
          If quickly Then .Refresh
        End If
      End With
    Else
      With flxASC
        If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
        If .Col <> oldCursorCol \ 2 Then .Col = oldCursorCol \ 2
        
        If .CellBackColor <> &H303030 Then
          .CellBackColor = &H303030
          .CellForeColor = &HC0C0&
        End If
      End With
      
      With flxHex
        If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
        If .Col <> oldCursorCol Then .Col = oldCursorCol
        
        If .CellBackColor <> &H505050 Then
          .CellBackColor = &H505050
          .CellForeColor = &HFFFF&
          
          If quickly Then .Refresh
        End If
      End With
    End If
    
    With flxOffset
      If .Row <> oldCursorRow + 1 Then .Row = oldCursorRow + 1
      
      If .CellForeColor <> &H8000000E Then
        .CellForeColor = &H8000000E
        .CellBackColor = &H80000010
        
        If quickly Then .Refresh
      End If
    End With
  End If
  
  If vFileSize - prjBlockSize > 0 Then
    LastVScroll = ((BaseOffset + PageOffset) / (vFileSize - prjBlockSize)) * 1000
    
   If LastVScroll > 999 Then LastVScroll = 999
   If LastVScroll < 1 Then LastVScroll = 1
  Else
    LastVScroll = 1
  End If
  
  VScroll1.Value = LastVScroll
  
  GenerateValues
  
  tmrCursorBlink.Enabled = True
End Sub

Private Sub LoadBlock()
  Dim TempAttrib As Long
  
  On Error GoTo goterror
  
  BlockSize = vFileSize - BaseOffset
  
  If BlockSize >= 48000 Then BlockSize = 48000
  If BlockSize <= 0 Then Exit Sub
  
  ReDim FileData(0 To BlockSize - 1)
  
  Get #1, BaseOffset + 1, FileData
  
  FileChange = False
  cmdSaveNow.Enabled = False
  cmdScrapNow.Enabled = False
  
  cmdViewDetails.Enabled = True
  cmdFind(0).Enabled = True
  cmdDelete.Enabled = True
  cmdInsert.Enabled = True
  
  Exit Sub
  
goterror:
  cmdViewDetails.Enabled = False
  cmdFind(0).Enabled = False
  cmdDelete.Enabled = False
  cmdInsert.Enabled = False
  
  MsgBox "There was a file access error!", vbOKOnly, "File Access Denied"
  
  Close #1
  
  TempAttrib = 0
    
  If chkArchive.Value = 1 Then TempAttrib = TempAttrib + vbArchive
  If chkRead.Value = 1 Then TempAttrib = TempAttrib + vbReadOnly
  If chkSystem.Value = 1 Then TempAttrib = TempAttrib + vbSystem
  If chkHidden.Value = 1 Then TempAttrib = TempAttrib + vbHidden
    
  If Not FileLock Then SetAttr vFilePath & vFileName, TempAttrib
  
  LastVScroll = 0
  BaseOffset = 0
  PageOffset = 0
  BlockSize = 0
  vFileSize = 0
  LineOffset = 0
  RowOffset = 0
  RowSubOffset = False
  FileLock = False
  
  VScroll1.Value = 0

  vFileName = ""
  
  lblLocked.Visible = True
  lblLocked.ToolTipText = "Read/Write Locked File"
End Sub

Private Sub SaveBlock()
  If FileChange Then
    If MsgBox("Save changes to file?", vbYesNo, "File Changed") = vbYes Then Put #1, BaseOffset + 1, FileData
  End If
  
  FileChange = False
  cmdScrapNow.Enabled = False
  cmdSaveNow.Enabled = False
End Sub

Private Sub tmrCursorBlink_Timer()
  Static blinkMode As Boolean
  
  blinkMode = Not blinkMode
  
  ShowCursor blinkMode, False
End Sub

Private Sub tmrFocus_Timer()
  On Error Resume Next
  
  VScroll1.SetFocus
End Sub

Private Sub VScroll1_Change()
  Dim Diff As Long
  
  If vFileSize <= 0 Then
    LastVScroll = 1
    VScroll1.Value = 1
    
    Exit Sub
  End If
  
  Diff = VScroll1.Value - LastVScroll
  
  If Diff = 0 Then Exit Sub
  
  LastVScroll = VScroll1.Value
  
  Select Case Diff
    Case 1
      Form_KeyDown vbKeyDown, 0
    Case 2
      Form_KeyDown vbKeyPageDown, 0
    Case -1
      Form_KeyDown vbKeyUp, 0
    Case -2
      Form_KeyDown vbKeyPageUp, 0
    Case Else
      Diff = vFileSize - prjBlockSize
      
      If Diff > 0 Then
        PageOffset = (((VScroll1.Value * (Diff / 1000)) + (prjLineWidth - 1)) \ prjLineWidth) * prjLineWidth
        
        If PageOffset + prjBlockSize >= vFileSize Then PageOffset = ((vFileSize + (prjLineWidth - 1)) \ prjLineWidth) * prjLineWidth - prjBlockSize
        If PageOffset < 0 Then PageOffset = 0
        
        If PageOffset > BaseOffset Then
          PageOffset = PageOffset - BaseOffset
          
          If PageOffset + prjBlockSize >= BlockSize Then
            PageOffset = PageOffset + BaseOffset
            
            SaveBlock
            
            BaseOffset = PageOffset - 24000
            If BaseOffset < 0 Then BaseOffset = 0
            
            PageOffset = PageOffset - BaseOffset
            
            LoadBlock
          End If
        Else
          SaveBlock
          
          BaseOffset = PageOffset - 24000
          If BaseOffset < 0 Then BaseOffset = 0
          
          PageOffset = PageOffset - BaseOffset
          
          LoadBlock
        End If
      Else
        LastVScroll = 1
        VScroll1.Value = 1
      End If
      
      UpdateDisplay
      
      If eMode Then
        flxASC.Refresh
      Else
        flxHex.Refresh
      End If
  End Select
End Sub

Public Sub GenerateValues()
  Dim vall As Long, vals As Long, Currentpos As Long
  Dim val1 As Long, val2 As Long, val3 As Long, val4 As Long
  
  On Error GoTo myerror
  
  If vFileName = "" Or vFileSize <= 0 Or FileLock Then
    lblValue(0).Caption = "--"
    lblValue(1).Caption = "--"
    lblValue(2).Caption = "--"
    lblValue(3).Caption = "--"
    lblValue(4).Caption = "--"
    lblValue(5).Caption = "--"
    lblValue(6).Caption = "--"
    lblValue(7).Caption = "--"
    lblValue(8).Caption = "--"
    lblValue(9).Caption = "--"
    
    Exit Sub
  End If
  
  Currentpos = PageOffset + (LineOffset * prjLineWidth) + RowOffset
  
  If aMode2 Then
    lblValue(9).Caption = Format(BaseOffset + Currentpos, "#,###,###,##0")
  Else
    lblValue(9).Caption = MakeHexLong(BaseOffset + Currentpos)
  End If
  
  val1 = FileData(Currentpos)
  
  lblValue(0).Caption = " " & val1
  lblValue(8).Caption = MakeHexShort(val1)
  
  If val1 > 127 Then
    lblValue(1).Caption = (val1 - 256) & " "
  Else
    lblValue(1).Caption = "+" & val1 & " "
  End If
  
  If Currentpos + 2 > BlockSize Then
    lblValue(2).Caption = "--"
    lblValue(3).Caption = "--"
    lblValue(4).Caption = "--"
    
    lblValue(5).Caption = "--"
    lblValue(6).Caption = "--"
    lblValue(7).Caption = "--"
  Else
    val2 = FileData(Currentpos + 1)
  
    vals = val2 * 256 + val1
    
    lblValue(2).Caption = Format(vals, " ##,##0")
    
    If val2 > 127 Then
      lblValue(3).Caption = Format(vals - 65536, "##,##0 ")
    Else
      lblValue(3).Caption = Format(vals, "+##,##0 ")
    End If
    
    lblValue(7).Caption = MakeHexShort(val2)
    
    If Currentpos + 4 > BlockSize Then
      lblValue(4).Caption = "--"
      
      lblValue(5).Caption = "--"
      lblValue(6).Caption = "--"
    Else
      val3 = FileData(Currentpos + 2)
      val4 = FileData(Currentpos + 3)
      lblValue(6).Caption = MakeHexShort(val3)
      lblValue(5).Caption = MakeHexShort(val4)
      
      If val4 > 127 Then
        val1 = val1 Xor &HFF
        val2 = val2 Xor &HFF
        val3 = val3 Xor &HFF
        val4 = val4 Xor &HFF
        
        If val1 = 255 And val2 = 255 And val3 = 255 And val4 = 127 Then
          lblValue(4).Caption = "-2,147,483,648"
        Else
          lblValue(4).Caption = Format((((val4 * 256 + val3) * 256) + val2) * 256 + val1 + 1, "-#,###,###,##0")
        End If
      Else
        lblValue(4).Caption = Format((((val4 * 256 + val3) * 256) + val2) * 256 + val1, "+#,###,###,##0")
      End If
    End If
  End If
  
  Exit Sub
  
myerror:
  MsgBox "got error", vbOKOnly
End Sub

Private Sub VScroll1_GotFocus()
  tmrFocus.Enabled = False
End Sub

Private Sub VScroll1_LostFocus()
  tmrFocus.Enabled = True
End Sub

Private Sub VScroll1_Scroll()
  VScroll1_Change
End Sub

Private Sub UpdateDisplay()
  Dim NumRows As Long, Row As Long, Column As Long
  Dim LastRow As Long, val As Long
  
  With flxHex
    NumRows = ((BlockSize - PageOffset) \ prjLineWidth) - 1
    LastRow = ((BlockSize - PageOffset) Mod prjLineWidth) - 1
    
    If NumRows >= prjLineHeight - 1 Then
      NumRows = prjLineHeight - 1
      LastRow = -1
    End If
    
    For Row = 0 To NumRows
      If aMode Then
        flxOffset.TextMatrix(Row + 1, 0) = Format(BaseOffset + PageOffset + Row * prjLineWidth, "#,###,###,##0")
      Else
        flxOffset.TextMatrix(Row + 1, 0) = MakeHexLong(BaseOffset + PageOffset + Row * prjLineWidth)
      End If
      
      For Column = 0 To (prjLineWidth - 1)
        val = FileData(PageOffset + Row * prjLineWidth + Column)
        
        .TextMatrix(Row + 1, Column * 2) = MakeHexShortH(val)
        .TextMatrix(Row + 1, Column * 2 + 1) = MakeHexShortL(val)
              
        If val >= 32 Then
          flxASC.TextMatrix(Row + 1, Column) = Chr$(val)
        Else
          flxASC.TextMatrix(Row + 1, Column) = "."
        End If
      Next Column
    Next Row
    
    If LastRow <> -1 Then
      If aMode Then
        flxOffset.TextMatrix(Row + 1, 0) = Format(BaseOffset + PageOffset + Row * prjLineWidth, "#,###,###,##0")
      Else
        flxOffset.TextMatrix(Row + 1, 0) = MakeHexLong(BaseOffset + PageOffset + Row * prjLineWidth)
      End If
      
      For Column = 0 To LastRow
        val = FileData(PageOffset + Row * prjLineWidth + Column)
        
        .TextMatrix(Row + 1, Column * 2) = MakeHexShortH(val)
        .TextMatrix(Row + 1, Column * 2 + 1) = MakeHexShortL(val)
        
        If val >= 32 Then
          flxASC.TextMatrix(Row + 1, Column) = Chr$(val)
        Else
          flxASC.TextMatrix(Row + 1, Column) = "."
        End If
      Next Column
      
      Do While Column < prjLineWidth
        .TextMatrix(Row + 1, Column * 2) = ""
        .TextMatrix(Row + 1, Column * 2 + 1) = ""
        
        flxASC.TextMatrix(Row + 1, Column) = ""
        
        Column = Column + 1
      Loop
      
      Row = Row + 1
    End If
    
    Do While Row < prjLineHeight
      flxOffset.TextMatrix(Row + 1, 0) = ""
      
      For Column = 0 To prjLineWidth - 1
        .TextMatrix(Row + 1, Column * 2) = ""
        .TextMatrix(Row + 1, Column * 2 + 1) = ""
        
        flxASC.TextMatrix(Row + 1, Column) = ""
      Next Column
      
      Row = Row + 1
    Loop
    
    If eMode Then
      .Refresh
    Else
      flxASC.Refresh
    End If
    
    flxOffset.Refresh
  End With
End Sub

