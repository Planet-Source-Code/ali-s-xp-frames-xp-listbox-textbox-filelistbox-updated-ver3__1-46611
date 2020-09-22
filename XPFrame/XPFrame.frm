VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WinXP Frames & TextBoxes"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin Project1.XPFrame XPFrame3 
      Height          =   2865
      Left            =   3225
      TabIndex        =   22
      Top             =   1650
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   5054
      Caption         =   "FileListBox correction (remove black border)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin VB.FileListBox File2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   150
         TabIndex        =   24
         Top             =   1800
         Width           =   3015
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   150
         TabIndex        =   23
         Top             =   450
         Width           =   3015
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Normal : (See the unwanted border)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   225
         Width           =   2610
      End
      Begin VB.Label Label4 
         Caption         =   "Just this code : File1.Appearance = 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   25
         Top             =   1575
         Width           =   2940
      End
   End
   Begin Project1.XPFrame XPFrame5 
      Height          =   1365
      Left            =   75
      TabIndex        =   10
      Top             =   4575
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2408
      Caption         =   "Normal Flat ListBox"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Align           =   2
      Begin VB.CommandButton Command1 
         Caption         =   "How to Use?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4425
         TabIndex        =   14
         Top             =   975
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   $"XPFrame.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   75
         TabIndex        =   15
         Top             =   300
         Width           =   6465
      End
   End
   Begin Project1.XPFrame XPFrame4 
      Height          =   2865
      Left            =   75
      TabIndex        =   6
      Top             =   1650
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   5054
      Caption         =   "Flat List Box (using XPListBox)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Align           =   2
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   1
         ItemData        =   "XPFrame.frx":0123
         Left            =   75
         List            =   "XPFrame.frx":0163
         TabIndex        =   16
         Top             =   1800
         Width           =   2880
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame4 
         Height          =   1065
         Left            =   75
         TabIndex        =   7
         Top             =   450
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1879
         Begin Project1.XPListBox XPListBox1 
            Height          =   975
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1720
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Index           =   0
               ItemData        =   "XPFrame.frx":024E
               Left            =   -15
               List            =   "XPFrame.frx":028E
               TabIndex        =   9
               Top             =   -15
               Width           =   2805
            End
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Normal ListBox :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   1575
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "XPListbox control :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   225
         Width           =   1335
      End
   End
   Begin Project1.XPFrame XPFrame2 
      Height          =   1515
      Left            =   3225
      TabIndex        =   1
      Top             =   75
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   2672
      Caption         =   "Multi Line"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   8421504
      FontName        =   "Tahoma"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin Project1.XPSimpleFrame XPSimpleFrame1 
         Height          =   315
         Left            =   225
         TabIndex        =   18
         Top             =   900
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   19
            Text            =   "Textbox with XPSimpleFrame"
            Top             =   45
            Width           =   2325
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   17
         Text            =   "Normal Textbox (using manifest)"
         Top             =   375
         Width           =   2415
      End
   End
   Begin Project1.XPFrame XPFrame1 
      Height          =   1515
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   2672
      Caption         =   "XP Frame"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1875
         TabIndex        =   13
         Top             =   900
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   12
         Top             =   600
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1875
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Font Italic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   675
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Font Bold"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Width           =   1215
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame2 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   1050
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   3
            Text            =   "XP Frame"
            Top             =   45
            Width           =   1500
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long

Private m_hMod As Long

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function


Private Sub Check1_Click()
    XPFrame1.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
    XPFrame1.FontItalic = Check2.Value
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Form_Initialize()
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControlsVB
End Sub

Private Sub Form_Load()
    File2.Appearance = 1
    If App.LogMode = 0 Then
        MsgBox "Compile to see all controls in xp style."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FreeLibrary m_hMod
End Sub

Private Sub Option1_Click(Index As Integer)
    XPFrame1.Align = Index
End Sub

Private Sub Option4_Click()
    'File1.Appearance = 0
    'File1.Refresh
    File1.Appearance = 1
    'File1.Refresh
End Sub

Private Sub Text2_Change()
    XPFrame1.Caption = Text2.Text
End Sub

