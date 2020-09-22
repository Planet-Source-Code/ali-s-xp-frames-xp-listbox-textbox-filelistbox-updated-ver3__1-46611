VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6870
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.XPFrame XPFrame3 
      Height          =   1440
      Left            =   150
      TabIndex        =   21
      Top             =   3600
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   2540
      Caption         =   "Thanks"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Align           =   2
      Begin VB.Label Label7 
         Height          =   990
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   2715
      End
   End
   Begin Project1.XPFrame XPFrame1 
      Height          =   5040
      Left            =   3375
      TabIndex        =   0
      Top             =   75
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   8890
      Caption         =   "Create XP Style ListBox"
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
      Align           =   2
      Begin Project1.XPSimpleFrame XPSimpleFrame4 
         Height          =   1140
         Left            =   1275
         TabIndex        =   11
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2011
         Begin Project1.XPListBox XPListBox1 
            Height          =   690
            Left            =   300
            TabIndex        =   14
            Top             =   225
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1217
         End
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame5 
         Height          =   1140
         Left            =   1275
         TabIndex        =   13
         Top             =   2100
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2011
         Begin Project1.XPListBox XPListBox2 
            Height          =   690
            Left            =   300
            TabIndex        =   15
            Top             =   150
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1217
            Begin VB.ListBox List1 
               Height          =   450
               ItemData        =   "Form2.frx":000C
               Left            =   75
               List            =   "Form2.frx":0013
               TabIndex        =   16
               Top             =   150
               Width           =   840
            End
         End
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame6 
         Height          =   1065
         Left            =   300
         TabIndex        =   17
         Top             =   3825
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   1879
         Begin Project1.XPListBox XPListBox3 
            Height          =   975
            Left            =   45
            TabIndex        =   18
            Top             =   45
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   1720
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               Height          =   1005
               ItemData        =   "Form2.frx":001E
               Left            =   -15
               List            =   "Form2.frx":0025
               TabIndex        =   19
               Top             =   -15
               Width           =   2730
            End
         End
      End
      Begin VB.Label Label6 
         Caption         =   "3. Set AutoSizeContained = True or       Resize the SimpleFrame."
         Height          =   420
         Left            =   225
         TabIndex        =   20
         Top             =   3375
         Width           =   2610
      End
      Begin VB.Label Label5 
         Caption         =   "2. Create a normal ListBox in XPListBox"
         Height          =   795
         Left            =   150
         TabIndex        =   12
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label Label4 
         Caption         =   "1. Create a Simple Frame and a XPListBox in it"
         Height          =   720
         Left            =   150
         TabIndex        =   10
         Top             =   375
         Width           =   2415
      End
   End
   Begin Project1.XPFrame XPFrame2 
      Height          =   3465
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   6112
      Caption         =   "Create XP Style TextBox"
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
      Align           =   2
      Begin Project1.XPSimpleFrame XPSimpleFrame1 
         Height          =   315
         Left            =   450
         TabIndex        =   3
         Top             =   675
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame2 
         Height          =   315
         Left            =   450
         TabIndex        =   5
         Top             =   1575
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   750
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   75
            Width           =   1290
         End
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame3 
         Height          =   315
         Left            =   450
         TabIndex        =   7
         Top             =   2700
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   45
            Width           =   900
         End
      End
      Begin VB.Label Label3 
         Caption         =   "3. Set AutoSizeContained = True or       Resize the SimpleFrame."
         Height          =   420
         Left            =   225
         TabIndex        =   6
         Top             =   2175
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2. Create a Texbox in the Frame"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   1200
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1. Create a Simple Frame"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   300
         Width           =   1770
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label7.Caption = "Thanks for visiting my program." & vbCrLf & _
                     "Please send your feedbacks and vote for me." & vbCrLf & _
                     "Email : ali6236@GameBox.net " & vbCrLf & _
                     "        or ali6236@yahoo.com"
End Sub
