VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AnswerCrack 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Phishbowler's Answer Machine Crack ©1998    "
   ClientHeight    =   7590
   ClientLeft      =   2550
   ClientTop       =   630
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "AnsCrack.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7590
   ScaleWidth      =   7305
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5520
      TabIndex        =   75
      Top             =   7920
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2400
      TabIndex        =   73
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   72
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dups not in list"
      Height          =   255
      Left            =   0
      TabIndex        =   71
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4920
      Top             =   7800
   End
   Begin VB.ListBox Attempt20lst 
      Height          =   1230
      Left            =   6600
      TabIndex        =   49
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt19lst 
      Height          =   1230
      Left            =   5880
      TabIndex        =   48
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt18lst 
      Height          =   1230
      Left            =   5160
      TabIndex        =   47
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt17lst 
      Height          =   1230
      Left            =   4440
      TabIndex        =   46
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt16lst 
      Height          =   1230
      Left            =   3720
      TabIndex        =   45
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt15lst 
      Height          =   1230
      Left            =   3000
      TabIndex        =   44
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt14lst 
      Height          =   1230
      Left            =   2280
      TabIndex        =   43
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt13lst 
      Height          =   1230
      Left            =   1560
      TabIndex        =   42
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt12lst 
      Height          =   1230
      Left            =   840
      TabIndex        =   41
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt11lst 
      Height          =   1230
      Left            =   120
      TabIndex        =   40
      Top             =   6240
      Width           =   615
   End
   Begin VB.ListBox Attempt10lst 
      Height          =   1230
      Left            =   6600
      TabIndex        =   39
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt9lst 
      Height          =   1230
      Left            =   5880
      TabIndex        =   38
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt8lst 
      Height          =   1230
      Left            =   5160
      TabIndex        =   37
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt7lst 
      Height          =   1230
      Left            =   4440
      TabIndex        =   36
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt6lst 
      Height          =   1230
      Left            =   3720
      TabIndex        =   35
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt5lst 
      Height          =   1230
      Left            =   3000
      TabIndex        =   34
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt4lst 
      Height          =   1230
      Left            =   2280
      TabIndex        =   33
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt3lst 
      Height          =   1230
      Left            =   1560
      TabIndex        =   32
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt2lst 
      Height          =   1230
      Left            =   840
      TabIndex        =   31
      Top             =   4560
      Width           =   615
   End
   Begin VB.ListBox Attempt1lst 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      TabIndex        =   30
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rem Dups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Attempt20txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Attempt 20"
      ToolTipText     =   "Click For Attempt 20"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Attempt19txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Attempt 19"
      ToolTipText     =   "Click For Attempt 19"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Attempt18txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Attempt 18"
      ToolTipText     =   "Click For Attempt 18"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Attempt17txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Attempt 17"
      ToolTipText     =   "Click For Attempt 17"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Attempt16txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Attempt 16"
      ToolTipText     =   "Click For Attempt 16"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Attempt15txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Attempt 15"
      ToolTipText     =   "Click For Attempt 15"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Attempt14txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Attempt 14"
      ToolTipText     =   "Click For Attempt 14"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Attempt13txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Attempt 13"
      ToolTipText     =   "Click For Attempt 13"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Attempt12txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Attempt 12"
      ToolTipText     =   "Click For Attempt 12"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Attempt11txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Attempt 11 "
      ToolTipText     =   "Click For Attempt 11"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Attempt10txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Attempt 10 "
      ToolTipText     =   "Click For Attempt 10"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Attempt9txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Attempt 9 "
      ToolTipText     =   "Click For Attempt 9"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Attempt8txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Attempt 8 "
      ToolTipText     =   "Click For Attempt 8"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Combotxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Combination"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Attempt7txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Attempt 7 "
      ToolTipText     =   "Click For Attempt 7"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "Joined Segments"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Attempt6txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Attempt 6 "
      ToolTipText     =   "Click For Attempt 6"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Attempt5txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Attempt 5 "
      ToolTipText     =   "Click For Attempt 5"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Attempt4txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Attempt 4 "
      ToolTipText     =   "Click For Attempt 4"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Attempt3txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Attempt 3 "
      ToolTipText     =   "Click For Attempt 3"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Attempt2txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Attempt 2 "
      ToolTipText     =   "Click for Attempt 2"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Attempt1txt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Attempt 1 "
      ToolTipText     =   "Click For Attempt1"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox DialBox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Play Tones"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox DigitTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Lists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Percentlbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0%"
      Height          =   255
      Left            =   6840
      TabIndex        =   74
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      Height          =   255
      Left            =   840
      TabIndex        =   70
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      Height          =   255
      Left            =   120
      TabIndex        =   69
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "18"
      Height          =   255
      Left            =   5160
      TabIndex        =   68
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "17"
      Height          =   255
      Left            =   4440
      TabIndex        =   67
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      Height          =   255
      Left            =   3720
      TabIndex        =   66
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      Height          =   255
      Left            =   3000
      TabIndex        =   65
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      Height          =   255
      Left            =   2280
      TabIndex        =   64
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      Height          =   255
      Left            =   1560
      TabIndex        =   63
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "20"
      Height          =   255
      Left            =   6600
      TabIndex        =   62
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      Height          =   255
      Left            =   6600
      TabIndex        =   61
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      Height          =   255
      Left            =   5880
      TabIndex        =   60
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      Height          =   255
      Left            =   5160
      TabIndex        =   59
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "19"
      Height          =   255
      Left            =   5880
      TabIndex        =   58
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      Height          =   255
      Left            =   4440
      TabIndex        =   57
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   255
      Left            =   3720
      TabIndex        =   56
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   255
      Left            =   3000
      TabIndex        =   55
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   255
      Left            =   2280
      TabIndex        =   54
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   255
      Left            =   1560
      TabIndex        =   53
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   255
      Left            =   840
      TabIndex        =   52
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      ToolTipText     =   "Double Click Each List to Remove Dups"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Individual Attempt Listed"
      Height          =   255
      Left            =   2640
      TabIndex        =   50
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Statuslbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
   Begin VB.Menu code 
      Caption         =   "Code Digits"
      Begin VB.Menu twodigit 
         Caption         =   "2 Digit Code"
      End
      Begin VB.Menu threedigit 
         Caption         =   "3 Digit Code"
      End
   End
   Begin VB.Menu section 
      Caption         =   "Sections"
      Begin VB.Menu moresec 
         Caption         =   "+10 Sections"
         Begin VB.Menu eleven 
            Caption         =   "11 Sections"
         End
         Begin VB.Menu twelvesec 
            Caption         =   "12 Sections"
         End
         Begin VB.Menu thirteensec 
            Caption         =   "13 Sections"
         End
         Begin VB.Menu fourteensec 
            Caption         =   "14 Sections"
         End
         Begin VB.Menu fifteensec 
            Caption         =   "15 Sections"
         End
         Begin VB.Menu sixteensec 
            Caption         =   "16 Sections"
         End
         Begin VB.Menu seventeensec 
            Caption         =   "17 Sections"
         End
         Begin VB.Menu eighteensec 
            Caption         =   "18 Sections"
         End
         Begin VB.Menu nineteensec 
            Caption         =   "19 Sections"
         End
         Begin VB.Menu twentysec 
            Caption         =   "20 Sections"
         End
      End
      Begin VB.Menu ThreeSec 
         Caption         =   "3 Sections"
      End
      Begin VB.Menu foursect 
         Caption         =   "4 Sections"
      End
      Begin VB.Menu fivesec 
         Caption         =   "5 Sections"
      End
      Begin VB.Menu sixsec 
         Caption         =   "6 Sections"
      End
      Begin VB.Menu sevensec 
         Caption         =   "7 Sections"
      End
      Begin VB.Menu eightsec 
         Caption         =   "8 Sections"
      End
      Begin VB.Menu ninesec 
         Caption         =   "9 Sections"
      End
      Begin VB.Menu tensec 
         Caption         =   "10 Sections"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "AnswerCrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()

End Sub

Private Sub Attempt10lst_DblClick()
If AnswerCrack.Attempt10lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt10lst)
End If
End Sub

Private Sub Attempt11lst_DblClick()
If AnswerCrack.Attempt11lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt11lst)
End If

End Sub

Private Sub Attempt12lst_DblClick()
If AnswerCrack.Attempt12lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt12lst)
End If

End Sub

Private Sub Attempt13lst_DblClick()
If AnswerCrack.Attempt13lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt13lst)
End If

End Sub

Private Sub Attempt14lst_DblClick()
If AnswerCrack.Attempt14lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt14lst)
End If

End Sub

Private Sub Attempt15lst_DblClick()
If AnswerCrack.Attempt15lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt15lst)
End If

End Sub

Private Sub Attempt16lst_DblClick()
If AnswerCrack.Attempt16lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt16lst)
End If

End Sub

Private Sub Attempt17lst_DblClick()
If AnswerCrack.Attempt17lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt17lst)
End If

End Sub

Private Sub Attempt18lst_DblClick()
If AnswerCrack.Attempt18lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt18lst)
End If

End Sub

Private Sub Attempt19lst_DblClick()
If AnswerCrack.Attempt19lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt19lst)
End If

End Sub

Private Sub Attempt1lst_DblClick()
If AnswerCrack.Attempt1lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt1lst)
End If
End Sub

Private Sub Attempt20lst_DblClick()
If AnswerCrack.Attempt20lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt20lst)
End If
End Sub



Private Sub Attempt2lst_DblClick()
If AnswerCrack.Attempt2lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt2lst)
End If
End Sub

Private Sub Attempt3lst_DblClick()
If AnswerCrack.Attempt3lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt3lst)
End If
End Sub

Private Sub Attempt4lst_DblClick()
If AnswerCrack.Attempt4lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt4lst)
End If

End Sub

Private Sub Attempt5lst_DblClick()
If AnswerCrack.Attempt4lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt4lst)
End If

End Sub

Private Sub Attempt6lst_DblClick()
If AnswerCrack.Attempt6lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt6lst)
End If

End Sub

Private Sub Attempt7lst_DblClick()
If AnswerCrack.Attempt7lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt7lst)
End If
End Sub

Private Sub Attempt8lst_DblClick()
If AnswerCrack.Attempt8lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt8lst)
End If

End Sub

Private Sub Attempt9lst_DblClick()
If AnswerCrack.Attempt9lst.ListCount > 0 Then
Call LISTKIllDuplicates(AnswerCrack.Attempt9lst)
End If

End Sub

Private Sub Command1_Click()
Dim percent As Integer
Dim Minutes As Integer
'Dim PauseAmt As Double
AnswerCrack.Timer1.Enabled = False
'Timer1.Enabled = True
Statuslbl = "Playing Tone Procedure"
Start = Timer                        'Time the procedure
t = 0
c = 0
'Start at 1st Chr
Amount = Len(AnswerCrack.DialBox)                  'Count Chrs in Text
If Amount = 0 Then Exit Sub
Static A
A = A + 1                      'Amount Times Clicked

Do
DoEvents
Bb$ = UCase(Mid$(AnswerCrack.DialBox, t + 1, 1))   'Go Through Each Chr
'Put Tones Here
WaveFile = Bb$
SoundName$ = App.Path & "\" & WaveFile & ".WAV"
wFlags% = SND_SYNC Or SND_NODEFAULT
'If Text2 = "" Then Text2 = 0
'PauseAmt = Text2

Call pause(0.05)
x% = sndPlaySound(SoundName$, wFlags%)

AnswerCrack.DigitTxt = Bb$
t = t + 1
'Clear Digit Box if Done
If t >= Len(AnswerCrack.DialBox) Then AnswerCrack.DigitTxt = ""
If AnswerCrack.threedigit.Checked = True Then
Cc$ = UCase(Mid$(AnswerCrack.DialBox, c + 1, 3))   'Find 3 Combos
TotalChrs = Len(AnswerCrack.DialBox)
percent = c / TotalChrs * 100
AnswerCrack.Percentlbl.Caption = percent & " %"
AnswerCrack.ProgressBar1.Value = percent

SpaceHere = InStr(Cc$, " ")
If SpaceHere Then GoTo Skip1Chr:
If Len(Cc$) = 1 Or Len(Cc$) = 2 Then GoTo Skip1Chr:
If AttemptNumber = 1 Then Attempt1lst.AddItem Cc$
If AttemptNumber = 2 Then Attempt2lst.AddItem Cc$
If AttemptNumber = 3 Then Attempt3lst.AddItem Cc$
If AttemptNumber = 4 Then Attempt4lst.AddItem Cc$
If AttemptNumber = 5 Then Attempt5lst.AddItem Cc$
If AttemptNumber = 6 Then Attempt6lst.AddItem Cc$
If AttemptNumber = 7 Then Attempt7lst.AddItem Cc$
If AttemptNumber = 8 Then Attempt8lst.AddItem Cc$
If AttemptNumber = 9 Then Attempt9lst.AddItem Cc$
If AttemptNumber = 10 Then Attempt10lst.AddItem Cc$
If AttemptNumber = 11 Then Attempt11lst.AddItem Cc$
If AttemptNumber = 12 Then Attempt12lst.AddItem Cc$
If AttemptNumber = 13 Then Attempt13lst.AddItem Cc$
If AttemptNumber = 14 Then Attempt14lst.AddItem Cc$
If AttemptNumber = 15 Then Attempt15lst.AddItem Cc$
If AttemptNumber = 16 Then Attempt16lst.AddItem Cc$
If AttemptNumber = 17 Then Attempt17lst.AddItem Cc$
If AttemptNumber = 18 Then Attempt18lst.AddItem Cc$
If AttemptNumber = 19 Then Attempt19lst.AddItem Cc$
If AttemptNumber = 20 Then Attempt20lst.AddItem Cc$
List1.AddItem Cc$
Skip1Chr:

End If

If AnswerCrack.twodigit.Checked = True Then
Cc$ = UCase(Mid$(AnswerCrack.DialBox, c + 1, 2))   'Find 2 Combos
TotalChrs = Len(AnswerCrack.DialBox)
percent = c / TotalChrs * 100
AnswerCrack.Percentlbl.Caption = percent & " %"
AnswerCrack.ProgressBar1.Value = percent
List1.AddItem Cc$
End If

c = c + 1

AnswerCrack.Combotxt = Cc$

If t = Amount Or t > Amount Then     'If Amount Filled..
 'Timer1.Enabled = False 'this just added!
 Finish = Timer
 Dim Done As Integer
 Done = (Finish - Start)
   If Done > 60 Then
   Minutes = Done / 60
   Seconds = (Minutes * 60) - Done
   Statuslbl = "Dialing took " & Minutes & "min. " & Seconds & "Seconds"
    Else
   Statuslbl = "Dialing took " & Done & " Seconds"
   End If
 
 GoTo Out                           'Then Done, Else

End If                               'Continue Loop
DoEvents
Loop
Out:
If percent = 100 Then AnswerCrack.Percentlbl = "0%"
AnswerCrack.Timer1.Enabled = True
'Remove Duplicates From Individual Lists

End Sub
Private Sub Command3_Click()
Static A
Start:
For x = 0 To List1.ListCount - 1
Itemz = List1.List(x)
A = A + 1
Text25 = Itemz & " " & x
If A = List2.ListCount Or A > List2.ListCount Then Exit Sub
If x = List1.ListCount Then GoTo Start:
List2.RemoveItem Itemz
Next x


End Sub

Private Sub Command2_Click()
'This is just a routine I used to check
'which duplicates I was missing in the
'list numbers 0-999, when testing each section.

Do
Item = List2.List(x)
If x >= List2.ListCount Then Exit Do
'If we have different Items then
If List1.List(x) <> List2.List(x) Then
Text1 = Text1 & " - " & List1.List(x)
End If
x = x + 1


Loop Until x = List1.ListCount


End Sub

Private Sub Command4_Click()
'For n = 0 To List1.ListCount - 1
'B = n - 1
'Text25 = "Memory: " & List1.List(B) & " " & B & "Number" & List1.List(n) & " " & n
'If B < 0 Then B = 0
'Memory = List1.List(B)
'If n = List1.ListCount Then Exit Sub
'If List1.List(n) = Memory Then List1.RemoveItem B
'Next n
Remember = AnswerCrack.Statuslbl.Caption
AnswerCrack.Statuslbl.Caption = "Removing Duplicates in List Please Wait..."
DoEvents
DuplicatesAmount = LISTKIllDuplicates(List1)
AnswerCrack.Statuslbl.Caption = Remember
msg = MsgBox(List1.ListCount & " in the list. " & DuplicatesAmount & " - Duplicates in List")
End Sub

Private Sub Command5_Click()
AnswerCrack.Attempt1lst.Clear
AnswerCrack.Attempt2lst.Clear
AnswerCrack.Attempt3lst.Clear
AnswerCrack.Attempt4lst.Clear
AnswerCrack.Attempt5lst.Clear
AnswerCrack.Attempt6lst.Clear
AnswerCrack.Attempt7lst.Clear
AnswerCrack.Attempt8lst.Clear
AnswerCrack.Attempt9lst.Clear
AnswerCrack.Attempt10lst.Clear
AnswerCrack.Attempt11lst.Clear
AnswerCrack.Attempt12lst.Clear
AnswerCrack.Attempt13lst.Clear
AnswerCrack.Attempt14lst.Clear
AnswerCrack.Attempt15lst.Clear
AnswerCrack.Attempt16lst.Clear
AnswerCrack.Attempt17lst.Clear
AnswerCrack.Attempt18lst.Clear
AnswerCrack.Attempt19lst.Clear
AnswerCrack.Attempt20lst.Clear
AnswerCrack.List1.Clear
End Sub

Private Sub Command6_Click()
WaveFile = AnswerCrack.DigitTxt
SoundName$ = "C:\SOUND\DTONE\" & WaveFile & "dtone.WAV"
wFlags% = SND_SYNC Or SND_NODEFAULT
x% = sndPlaySound(SoundName$, wFlags%)

End Sub

Private Sub DialBox_DblClick()

Command1_Click
End Sub

Private Sub eighteensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "18 Sections Approx - 7 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt16txt = ""
AnswerCrack.Attempt17txt = ""
AnswerCrack.Attempt18txt = ""
AnswerCrack.Attempt19txt = ""

AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = True
AnswerCrack.Attempt16txt.Visible = True
AnswerCrack.Attempt17txt.Visible = True
AnswerCrack.Attempt18txt.Visible = True
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False


AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "902300310320330340350360370380390300300410420430440450460470480490"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "400500510520530540550560570580590506006106206306406506606706806906"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "070071072073074075076077078079070800810820830840850860870880890809"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "0091092093094095096097098099099111113131132133211212221222231232"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "233311312313321322323331332333411412413414421422423424431432433434"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "441442443444511512513514515521522523524525531532533534535541542543"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "544545551552553554555611612613614615616621622623624625626631632633"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "634635636641642643644645646651652653654655656661662663664665666711"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "712713714715716717721722723724725726727731732733734735736737741742"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "743744745746747751752753754755756757761762763764765766767771772773"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "774775776777811812813814815816817818821822823824825826827828831832"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "833834835836837838841842843844845846847848851852853854855856857858"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "861862863864865866867868871872873874875876877878881882883884885886"
AnswerCrack.Attempt14txt.Enabled = True
AnswerCrack.Attempt15txt = "887888911912913914915916917918919921922923924925926927928929931932"
AnswerCrack.Attempt15txt.Enabled = True
AnswerCrack.Attempt16txt = "933934935936937938939941942943944945946947948949951952953954955956"
AnswerCrack.Attempt16txt.Enabled = True
AnswerCrack.Attempt17txt = "957958959961962963964965966967968969971972973974975976977978979981"
AnswerCrack.Attempt17txt.Enabled = True
AnswerCrack.Attempt18txt = "982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt18txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr
RightBox15a = Right(AnswerCrack.Attempt15txt, 1)  ' 15Box 3 First 2 Chr
RightBox15b = Right(AnswerCrack.Attempt15txt, 2)  ' 15Box 3 First 2 Chr
RightBox16a = Right(AnswerCrack.Attempt16txt, 1)  ' 15Box 3 First 2 Chr
RightBox16b = Right(AnswerCrack.Attempt16txt, 2)  ' 15Box 3 First 2 Chr
RightBox17a = Right(AnswerCrack.Attempt17txt, 1)  ' 15Box 3 First 2 Chr
RightBox17b = Right(AnswerCrack.Attempt17txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)
Box26a = Left(AnswerCrack.Attempt15txt, 1)
Box26b = Left(AnswerCrack.Attempt15txt, 2)
Box27a = Left(AnswerCrack.Attempt16txt, 1)
Box27b = Left(AnswerCrack.Attempt16txt, 2)
Box28a = Left(AnswerCrack.Attempt17txt, 1)
Box28b = Left(AnswerCrack.Attempt17txt, 2)
Box29a = Left(AnswerCrack.Attempt18txt, 1)
Box29b = Left(AnswerCrack.Attempt18txt, 2)


                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a & RightBox14a & Box26b & RightBox14b & Box26a & RightBox15a & Box27b & RightBox15b & Box27a & RightBox16a & Box28b & RightBox16b & Box28a & RightBox17a & Box29b & RightBox17b & Box29a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt18txt = AnswerCrack.Attempt18txt & Extras



End Sub

Private Sub eightsec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "8 Sections Approx - 17 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt16txt = ""
AnswerCrack.Attempt17txt = ""
AnswerCrack.Attempt18txt = ""
AnswerCrack.Attempt19txt = ""
AnswerCrack.Attempt20txt = ""


AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390300410420430440450460470480490400500510520530540"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "550560570580590506006106206306406506606706806906070071072073074075076077078079070800810820830840850860870880890809009109209309409509609709809909911"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "111313113213321121222122223123223331131231332132232333133233341141241341442142242342443143243343444144244344451151251351451552152252352452553153253"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "353453554154254354454555155255355455561161261361461561662162262362462562663163263363463563664164264364464564665165265365465565666166266366466566671"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "171271371471571671772172272372472572672773173273373473573673774174274374474574674775175275375475575675776176276376476576676777177277377477577677781"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "181281381481581681781882182282382482582682782883183283383483583683783884184284384484584684784885185285385485585685785886186286386486586686786887187"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "287387487587687787888188288388488588688788891191291391491591691791891992192292392492592692792892993193293393493593693793893994194294394494594694794"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "8949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt8txt.Enabled = True


AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

'Take the first character all the way on the right
'Take the first two all the way on the right
'30203  3 and 03

Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  '7 First 1 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  '7 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------

'First character on left = a
'First two characters on left = b

Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
                      
        '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a '& Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt8txt = AnswerCrack.Attempt8txt & Extras

End Sub

Private Sub eleven_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "11 Sections Approx - 12 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False


AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390300410"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "420430440450460470480490400500510520530540550560570580590506006106206306406506606706806906070071072073074"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "075076077078079070800810820830840850860870880890809009109209309409509609709809909911111313113213321121222"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "122223123223331131231332132232333133233341141241341442142242342443143243343444144244344451151251351451552"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "152252352452553153253353453554154254354454555155255355455561161261361461561662162262362462562663163263363"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "463563664164264364464564665165265365465565666166266366466566671171271371471571671772172272372472572672773"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "173273373473573673774174274374474574674775175275375475575675776176276376476576676777177277377477577677781"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "181281381481581681781882182282382482582682782883183283383483583683783884184284384484584684784885185285385"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "485585685785886186286386486586686786887187287387487587687787888188288388488588688788891191291391491591691"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "79189199219229239249259269279289299319329339349359369379389399419429439449459469479489499519529539549559"
AnswerCrack.Attempt10txt.Enabled = True
'AnswerCrack.Attempt9txt = "956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"

AnswerCrack.Attempt11txt = "56957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt11txt.Enabled = True

'AnswerCrack.Attempt11txt = ""
'AnswerCrack.Attempt12txt = ""
'AnswerCrack.Attempt13txt = ""
'AnswerCrack.Attempt14txt = ""
'AnswerCrack.Attempt15txt = ""
'AnswerCrack.Attempt16txt = ""
'AnswerCrack.Attempt17txt = ""
'AnswerCrack.Attempt18txt = ""
'AnswerCrack.Attempt19txt = ""
Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
                      
                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a '& Box11a & Box23b & Box11b & Box23a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt11txt = AnswerCrack.Attempt11txt & Extras

End Sub

Private Sub fifteensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "14 Sections Approx - 9 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = True
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

AnswerCrack.Attempt1txt = "0001001101201301401501601701801901020120021022023024025026027028029023003103203"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "3034035036037038039030030041042043044045046047048049040050051052053054055056057"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "0580590506006106206306406506606706806906070071072073074075076077078079070800810"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "82083084085086087088089080900910920930940950960970980990991111131311321332112"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "1222122223123223331131231332132232333133233341141241341442142242342443143243343"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "4441442443444511512513514515521522523524525531532533534535541542543544545551552"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "5535545556116126136146156166216226236246256266316326336346356366416426436446456"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "4665165265365465565666166266366466566671171271371471571671772172272372472572672"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "7731732733734735736737741742743744745746747751752753754755756757761762763764765"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "7667677717727737747757767778118128138148158168178188218228238248258268278288318"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "3283383483583683783884184284384484584684784885185285385485585685785886186286386"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "4865866867868871872873874875876877878881882883884885886887888911912913914915916"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "9179189199219229239249259269279289299319329339349359369379389399419429439449459"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "4694794894995195295395495595695795895996196296396496596696796896997197297397497"
AnswerCrack.Attempt14txt.Enabled = True
AnswerCrack.Attempt15txt = "5976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt15txt.Enabled = True



Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)
Box26a = Left(AnswerCrack.Attempt15txt, 1)
Box26b = Left(AnswerCrack.Attempt15txt, 2)

                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a & RightBox14a & Box26b & RightBox14b & Box26a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt15txt = AnswerCrack.Attempt15txt & Extras



End Sub

Private Sub fivesec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "5 Sections Approx - 26 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390300410420430440450460470480490400500510520530540550560570580590506006106206306406506606706806906070071072073074075076077078079070800810"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "820830840850860870880890809009109209309409509609709809909911111313113213321121222122223123223331131231332132232333133233341141241341442142242342443143243343444144244344451151251351451552152252352452553153253353453554154254354454555155"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "255355455561161261361461561662162262362462562663163263363463563664164264364464564665165265365465565666166266366466566671171271371471571671772172272372472572672773173273373473573673774174274374474574674775175275375475575675776176276376"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "476576676777177277377477577677781181281381481581681781882182282382482582682782883183283383483583683783884184284384484584684784885185285385485585685785886186286386486586686786887187287387487587687787888188288388488588688788891191291391"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "4915916917918919921922923924925926927928929931932933934935936937938939941942943944945946947948949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt5txt.Enabled = True



AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
           
           '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a    '& Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box15b & Box7b & Box15a & Box8a & Box20b & Box8b & Box20a & Box8a & Box19b & Box8b & Box19a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt5txt = AnswerCrack.Attempt5txt & Extras
                   

End Sub

Private Sub Form_Load()
'Created By Phishbowler
'Visit me: http://come.to/phishbowler

'Load the program's Icon at Runtime
TheIcon = App.Path & "\Answer Crack.ico"
AnswerCrack.Icon = LoadPicture(TheIcon)

msg = MsgBox("This Program Is For Educational Use Only. Do NOT use this program for ILLEGAL purposes. If any damages or legal matters arise because you use this program entitled Answer Machine Crack, the maker of this program will NOT be held accountable or liable in any way whatsoever. Use this at YOUR OWN RISK. If you do not agree to this, EXIT THE PROGRAM WHEN IT LOADS IMMEDIATELY", vbOKOnly, "DISCLAIMER - READ THIS OR DONT USE THIS PROGRAM")

'Add 000-999 to a List
For A = 0 To 9
List2.AddItem 0 & 0 & A
Next A
For B = 10 To 99
List2.AddItem 0 & B
Next B
For x = 100 To 999
List2.AddItem x
Next x

'Hide all the segments
AnswerCrack.Attempt1txt.Visible = False
AnswerCrack.Attempt2txt.Visible = False
AnswerCrack.Attempt3txt.Visible = False
AnswerCrack.Attempt4txt.Visible = False
AnswerCrack.Attempt5txt.Visible = False
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False


'FOR COMMENTS FOR SECTION CODES
'SEE THE 4th Section right below
'this comment, all other sections
'are the same.

End Sub

Private Sub foursect_Click()
'Display Four Sections

AnswerCrack.List1.Clear
Statuslbl = "4 Sections Approx - 43 Intervals"
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call pause(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call pause(2): Statuslbl = "": Exit Sub
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
                                   '00010011012013014015016017018019010201200210220230240250260270280290230031032033034035036037038039030041042043044045046047048049040050051052053054055056057058059050600610620630640650660670680690607007107207307407507607707807907080081082083084085086087088089080900910920930940950960970980990991111131311321332112122212222312322333113123133213223233313323334114124134144214224234244314324334344414424434445
AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250260270280290230031032033034035036037038039030041042043044045046047048049040050051052053054055056057058059050600610620630640650660670680690607007107207307407507607707807907080081082083084085086087088089080900910920930940950960970980990991111131311321332112122212222312322333113123133213223233313323334114124134144214224234244314324"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "940950960970980991111131311321332112122212222312322333113123133213223233313323334114124134144214224234244314324334344414424434445115125135145155215225235245255315325335345355415425435445455515525535545556116126136146156166216226236246256266316326336346356366416426436446456466516"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "5265365465565666166266366466566671171271371471571671772172272372472572672773173273373473573673774174274374474574674775175275375475575675776176276376476576676777177277377477577677781181281381481581681781882182282382482582682782883183283383483583683783884184284384484584684784885185"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "2853854855856857858861862863864865866867868871872873874875876877878881882883884885886887888911912913914915916917918919921922923924925926927928929931932933934935936937938939941942943944945946947948949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt4txt.Enabled = True

AnswerCrack.Attempt5txt.Visible = False
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False

'Below is any codes that may be cut off from segmenting

Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
                      
'The extras are the divided numbers you lose when you segment
                     
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a    '& Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box15b & Box7b & Box15a & Box8a & Box20b & Box8b & Box20a & Box8a & Box19b & Box8b & Box19a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a

'This was just for testing purposes
Text10 = Extras

'Add On Extras to Last Box
AnswerCrack.Attempt4txt = AnswerCrack.Attempt4txt & Extras

'The last box will end up taking longer because of extras
'But at least all numbers are in fact present.

End Sub

Private Sub fourteensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "14 Sections Approx - 10 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

AnswerCrack.Attempt1txt = "0001001101201301401501601701801901020120021022023024025026027028029023003103203303403"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "5036037038039030030041042043044045046047048049040050051052053054055056057058059050600"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "6106206306406506606706806906070071072073074075076077078079070800810820830840850860870"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "88089080900910920930940950960970980990991111131311321332112122212222312322333113123"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "1332132232333133233341141241341442142242342443143243343444144244344451151251351451552"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "1522523524525531532533534535541542543544545551552553554555611612613614615616621622623"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "6246256266316326336346356366416426436446456466516526536546556566616626636646656667117"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "1271371471571671772172272372472572672773173273373473573673774174274374474574674775175"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "2753754755756757761762763764765766767771772773774775776777811812813814815816817818821"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "8228238248258268278288318328338348358368378388418428438448458468478488518528538548558"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "5685785886186286386486586686786887187287387487587687787888188288388488588688788891191"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "2913914915916917918919921922923924925926927928929931932933934935936937938939941942943"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "9449459469479489499519529539549559569579589599619629639649659669679689699719729739749"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "75976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt14txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)

                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt14txt = AnswerCrack.Attempt14txt & Extras



End Sub

Private Sub help_Click()
Helpfrm.Show
End Sub

Private Sub List1_DblClick()
List1.Clear
End Sub

Private Sub ninesec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "9 Sections Approx - 15 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False



AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250260270280290230031032033034035036037038039030041042043044045046047048049040"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "05005105205305405505605705805905060061062063064065066067068069060700710720730740750760770780790708008108208308408508608708808908090"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "09109209309409509609709809909911111313113213321121222122223123223331131231332132232333133233341141241341442142242342443143243343444"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "14424434445115125135145155215225235245255315325335345355415425435445455515525535545556116126136146156166216226236246256266316326336"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "34635636641642643644645646651652653654655656661662663664665666711712713714715716717721722723724725726727731732733734735736737741742"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "74374474574674775175275375475575675776176276376476576676777177277377477577677781181281381481581681781882182282382482582682782883183"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "28338348358368378388418428438448458468478488518528538548558568578588618628638648658668678688718728738748758768778788818828838848858"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "86887888911912913914915916917918919921922923924925926927928929931932933934935936937938939941942943944945946947948949951952953954955"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt9txt.Enabled = True

Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
                      
        '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a '& Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt9txt = AnswerCrack.Attempt9txt & Extras


End Sub

Private Sub nineteensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "18 Sections Approx - 7 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt16txt = ""
AnswerCrack.Attempt17txt = ""
AnswerCrack.Attempt18txt = ""
AnswerCrack.Attempt19txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = True
AnswerCrack.Attempt16txt.Visible = True
AnswerCrack.Attempt17txt.Visible = True
AnswerCrack.Attempt18txt.Visible = True
AnswerCrack.Attempt19txt.Visible = True
AnswerCrack.Attempt20txt.Visible = False


AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250260270"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "28029023003103203303403503603703803903003004104204304404504604"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "70480490400500510520530540550560570580590506006106206306406506"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "60670680690607007107207307407507607707807907080081082083084085"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "086087088089080900910920930940950960970980990991111131311321"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "33211212221222231232233311312313321322323331332333411412413414"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "42142242342443143243343444144244344451151251351451552152252352"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "45255315325335345355415425435445455515525535545556116126136146"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "15616621622623624625626631632633634635636641642643644645646651"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "65265365465565666166266366466566671171271371471571671772172272"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "37247257267277317327337347357367377417427437447457467477517527"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "53754755756757761762763764765766767771772773774775776777811812"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "81381481581681781882182282382482582682782883183283383483583683"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "78388418428438448458468478488518528538548558568578588618628638"
AnswerCrack.Attempt14txt.Enabled = True
AnswerCrack.Attempt15txt = "64865866867868871872873874875876877878881882883884885886887888"
AnswerCrack.Attempt15txt.Enabled = True
AnswerCrack.Attempt16txt = "91191291391491591691791891992192292392492592692792892993193293"
AnswerCrack.Attempt16txt.Enabled = True
AnswerCrack.Attempt17txt = "39349359369379389399419429439449459469479489499519529539549559"
AnswerCrack.Attempt17txt.Enabled = True
AnswerCrack.Attempt18txt = "56957958959961962963964965966967968969971972973974975976977978"
AnswerCrack.Attempt18txt.Enabled = True
AnswerCrack.Attempt19txt = "979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt19txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr
RightBox15a = Right(AnswerCrack.Attempt15txt, 1)  ' 15Box 3 First 2 Chr
RightBox15b = Right(AnswerCrack.Attempt15txt, 2)  ' 15Box 3 First 2 Chr
RightBox16a = Right(AnswerCrack.Attempt16txt, 1)  ' 15Box 3 First 2 Chr
RightBox16b = Right(AnswerCrack.Attempt16txt, 2)  ' 15Box 3 First 2 Chr
RightBox17a = Right(AnswerCrack.Attempt17txt, 1)  ' 15Box 3 First 2 Chr
RightBox17b = Right(AnswerCrack.Attempt17txt, 2)  ' 15Box 3 First 2 Chr
RightBox18a = Right(AnswerCrack.Attempt18txt, 1)  ' 15Box 3 First 2 Chr
RightBox18b = Right(AnswerCrack.Attempt18txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)
Box26a = Left(AnswerCrack.Attempt15txt, 1)
Box26b = Left(AnswerCrack.Attempt15txt, 2)
Box27a = Left(AnswerCrack.Attempt16txt, 1)
Box27b = Left(AnswerCrack.Attempt16txt, 2)
Box28a = Left(AnswerCrack.Attempt17txt, 1)
Box28b = Left(AnswerCrack.Attempt17txt, 2)
Box29a = Left(AnswerCrack.Attempt18txt, 1)
Box29b = Left(AnswerCrack.Attempt18txt, 2)
Box30a = Left(AnswerCrack.Attempt19txt, 1)
Box30b = Left(AnswerCrack.Attempt19txt, 2)


                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a & RightBox14a & Box26b & RightBox14b & Box26a & RightBox15a & Box27b & RightBox15b & Box27a & RightBox16a & Box28b & RightBox16b & Box28a & RightBox17a & Box29b & RightBox17b & Box29a & RightBox18a & Box30b & RightBox18b & Box30a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt19txt = AnswerCrack.Attempt19txt & Extras


End Sub

Private Sub sevensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "7 Sections Approx - 19 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390300410420430440450460470480490400500510520530540550560570580590506006"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "106206306406506606706806906070071072073074075076077078079070800810820830840850860870880890809009109209309409509609709809909911111313113213321121222122223123223331131231"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "332132232333133233341141241341442142242342443143243343444144244344451151251351451552152252352452553153253353453554154254354454555155255355455561161261361461561662162262"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "362462562663163263363463563664164264364464564665165265365465565666166266366466566671171271371471571671772172272372472572672773173273373473573673774174274374474574674775"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "175275375475575675776176276376476576676777177277377477577677781181281381481581681781882182282382482582682782883183283383483583683783884184284384484584684784885185285385"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "485585685785886186286386486586686786887187287387487587687787888188288388488588688788891191291391491591691791891992192292392492592692792892993193293393493593693793893994"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "1942943944945946947948949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt7txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
        '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & " " & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a    '& Box7a & Box15b & Box7b & Box15a & Box8a & Box20b & Box8b & Box20a & Box8a & Box19b & Box8b & Box19a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt7txt = AnswerCrack.Attempt7txt & Extras
                      

End Sub

Private Sub seventeensec_Click()
AnswerCrack.List1.Clear
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "17 Sections Approx - 8 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt16txt = ""
AnswerCrack.Attempt17txt = ""
AnswerCrack.Attempt18txt = ""
AnswerCrack.Attempt19txt = ""


AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = True
AnswerCrack.Attempt16txt.Visible = True
AnswerCrack.Attempt17txt.Visible = True
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False


AnswerCrack.Attempt1txt = "0001001101201301401501601701801901020120021022023024025026027028029023"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "0031032033034035036037038039030030041042043044045046047048049040050051"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "0520530540550560570580590506006106206306406506606706806906070071072073"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "0740750760770780790708008108208308408508608708808908090091092093094095"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "09609709809909911111313113213321121222122223123223331131231332132232"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "3331332333411412413414421422423424431432433434441442443444511512513514"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "5155215225235245255315325335345355415425435445455515525535545556116126"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "1361461561662162262362462562663163263363463563664164264364464564665165"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "2653654655656661662663664665666711712713714715716717721722723724725726"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "7277317327337347357367377417427437447457467477517527537547557567577617"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "6276376476576676777177277377477577677781181281381481581681781882182282"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "3824825826827828831832833834835836837838841842843844845846847848851852"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "8538548558568578588618628638648658668678688718728738748758768778788818"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "8288388488588688788891191291391491591691791891992192292392492592692792"
AnswerCrack.Attempt14txt.Enabled = True
AnswerCrack.Attempt15txt = "8929931932933934935936937938939941942943944945946947948949951952953954"
AnswerCrack.Attempt15txt.Enabled = True
AnswerCrack.Attempt16txt = "9559569579589599619629639649659669679689699719729739749759769779789799"
AnswerCrack.Attempt16txt.Enabled = True
AnswerCrack.Attempt17txt = "81982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt17txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr
RightBox15a = Right(AnswerCrack.Attempt15txt, 1)  ' 15Box 3 First 2 Chr
RightBox15b = Right(AnswerCrack.Attempt15txt, 2)  ' 15Box 3 First 2 Chr
RightBox16a = Right(AnswerCrack.Attempt16txt, 1)  ' 15Box 3 First 2 Chr
RightBox16b = Right(AnswerCrack.Attempt16txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)
Box26a = Left(AnswerCrack.Attempt15txt, 1)
Box26b = Left(AnswerCrack.Attempt15txt, 2)
Box27a = Left(AnswerCrack.Attempt16txt, 1)
Box27b = Left(AnswerCrack.Attempt16txt, 2)
Box28a = Left(AnswerCrack.Attempt17txt, 1)
Box28b = Left(AnswerCrack.Attempt17txt, 2)


                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a & RightBox14a & Box26b & RightBox14b & Box26a & RightBox15a & Box27b & RightBox15b & Box27a & RightBox16a & Box28b & RightBox16b & Box28a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt17txt = AnswerCrack.Attempt17txt & Extras



End Sub

Private Sub sixsec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "6 Sections Approx - ? Second Intervals"
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""

AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250260270280290230031032033034035036037038039030041042043044045046047048049040050051052053054055056057058059050600610620630640650660670680690607007"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "107207307407507607707807907080081082083084085086087088089080900910920930940950960970980990991111131311321332112122212222312322333113123133213223233313323334114124134144214224234244314324334344414424"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "4344451151251351451552152252352452553153253353453554154254354454555155255355455561161261361461561662162262362462562663163263363463563664164264364464564665165265365465565666166266366466566671171"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "27137147157167177217227237247257267277317327337347357367377417427437447457467477517527537547557567577617627637647657667677717727737747757767778118128138148158168178188218228238248258268278288"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "3183283383483583683783884184284384484584684784885185285385485585685785886186286386486586686786887187287387487587687787888188288388488588688788891191291391491591691791891992192292392492592692792"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "8929931932933934935936937938939941942943944945946947948949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt7txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
       '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a    '& Box6a & Box18b & Box6b & Box18a & Box7a & Box15b & Box7b & Box15a & Box8a & Box20b & Box8b & Box20a & Box8a & Box19b & Box8b & Box19a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt6txt = AnswerCrack.Attempt6txt & Extras
End Sub

Private Sub sixteensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "14 Sections Approx - 8 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt16txt = ""
AnswerCrack.Attempt17txt = ""
AnswerCrack.Attempt18txt = ""
AnswerCrack.Attempt19txt = ""

AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = True
AnswerCrack.Attempt16txt.Visible = True
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False


AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250260270280290230031"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "03203303403503603703803903003004104204304404504604704804904005005105205305"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "40550560570580590506006106206306406506606706806906070071072073074075076077"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "07807907080081082083084085086087088089080900910920930940950960970980990991"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "111131311321332112122212222312322333113123133213223233313323334114124134"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "14421422423424431432433434441442443444511512513514515521522523524525531532"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "53353453554154254354454555155255355455561161261361461561662162262362462562"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "66316326336346356366416426436446456466516526536546556566616626636646656667"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "11712713714715716717721722723724725726727731732733734735736737741742743744"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "74574674775175275375475575675776176276376476576676777177277377477577677781"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "18128138148158168178188218228238248258268278288318328338348358368378388418"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "42843844845846847848851852853854855856857858861862863864865866867868871872"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "87387487587687787888188288388488588688788891191291391491591691791891992192"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "29239249259269279289299319329339349359369379389399419429439449459469479489"
AnswerCrack.Attempt14txt.Enabled = True
AnswerCrack.Attempt15txt = "49951952953954955956957958959961962963964965966967968969971972973974975976"
AnswerCrack.Attempt15txt.Enabled = True
AnswerCrack.Attempt16txt = "977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt16txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr
RightBox15a = Right(AnswerCrack.Attempt15txt, 1)  ' 15Box 3 First 2 Chr
RightBox15b = Right(AnswerCrack.Attempt15txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)
Box26a = Left(AnswerCrack.Attempt15txt, 1)
Box26b = Left(AnswerCrack.Attempt15txt, 2)
Box27a = Left(AnswerCrack.Attempt16txt, 1)
Box27b = Left(AnswerCrack.Attempt16txt, 2)

                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a & RightBox14a & Box26b & RightBox14b & Box26a & RightBox15a & Box27b & RightBox15b & Box27a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt16txt = AnswerCrack.Attempt16txt & Extras



End Sub

Private Sub tensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "10 Sections Approx - 13 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False


AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390300410420430440450"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "460470480490400500510520530540550560570580590506006106206306406506606706806906070071072073074075076077078079070800810"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "820830840850860870880890809009109209309409509609709809909911111313113213321121222122223123223331131231332132232333133"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "233341141241341442142242342443143243343444144244344451151251351451552152252352452553153253353453554154254354454555155"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "255355455561161261361461561662162262362462562663163263363463563664164264364464564665165265365465565666166266366466566"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "671171271371471571671772172272372472572672773173273373473573673774174274374474574674775175275375475575675776176276376"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "476576676777177277377477577677781181281381481581681781882182282382482582682782883183283383483583683783884184284384484"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "584684784885185285385485585685785886186286386486586686786887187287387487587687787888188288388488588688788891191291391"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "491591691791891992192292392492592692792892993193293393493593693793893994194294394494594694794894995195295395495595695"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "7958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt10txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr

                      
        '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a '& Box11a & Box23b & Box11b & Box23a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt10txt = AnswerCrack.Attempt10txt & Extras

End Sub

Private Sub Text1_Change()
'Dim S As Integer
'If AnswerCrack.DigitTxt = "" Then Exit Sub
'On Error GoTo Shit2:
'WaveFile = AnswerCrack.DigitTxt
'SoundName$ = "C:\SOUND\REAL\" & WaveFile & ".WAV"
'wflags% = SND_SYNC Or SND_NODEFAULT
'x% = sndplaysound(SoundName$, wflags%)
'AnswerCrack.DigitTxt = ""
'Shit2:
'Resume 2
'2 :

End Sub

Private Sub Attempt7txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt7txt
Statuslbl = "Crack Attempt 7"
AnswerCrack.Attempt7txt.Enabled = False
AttemptNumber = 7
Command1_Click
Command1.Caption = "Play Attempt 7"
End Sub

Private Sub Attempt8txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt8txt
Statuslbl = "Crack Attempt 8"
AnswerCrack.Attempt8txt.Enabled = False
AttemptNumber = 8
Command1_Click
Command1.Caption = "Play Attempt 8"
End Sub

Private Sub Attempt9txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt9txt
Statuslbl = "Crack Attempt 9"
AnswerCrack.Attempt9txt.Enabled = False
AttemptNumber = 9
Command1_Click
Command1.Caption = "Play Attempt 9"
End Sub

Private Sub Attempt10txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt10txt
Statuslbl = "Crack Attempt 10"
AnswerCrack.Attempt10txt.Enabled = False
AttemptNumber = 10
Command1_Click
Command1.Caption = "Play Attempt 10"
End Sub

Private Sub Attempt11txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt11txt
Statuslbl = "Crack Attempt 11"
AnswerCrack.Attempt11txt.Enabled = False
AttemptNumber = 11
Command1_Click
Command1.Caption = "Play Attempt 11"
End Sub

Private Sub Attempt12txt_Click()
Statuslbl = "Crack Attempt 12"
AnswerCrack.DialBox = AnswerCrack.Attempt12txt
AnswerCrack.Attempt12txt.Enabled = False
AttemptNumber = 12
Command1_Click
Command1.Caption = "Play Attempt 12"
End Sub

Private Sub Attempt13txt_Click()
Statuslbl = "Crack Attempt 13"
AnswerCrack.DialBox = AnswerCrack.Attempt13txt
AnswerCrack.Attempt13txt.Enabled = False
AttemptNumber = 13
Command1_Click
Command1.Caption = "Play Attempt 13"
End Sub

Private Sub Attempt14txt_Click()
Statuslbl = "Crack Attempt 14"
AnswerCrack.DialBox = AnswerCrack.Attempt14txt
AnswerCrack.Attempt14txt.Enabled = False
AttemptNumber = 14
Command1_Click
Command1.Caption = "Play Attempt 14"
End Sub

Private Sub Attempt15txt_Click()
Statuslbl = "Crack Attempt 15"
AnswerCrack.DialBox = AnswerCrack.Attempt15txt
AnswerCrack.Attempt15txt.Enabled = False
AttemptNumber = 15
Command1_Click
Command1.Caption = "Play Attempt 15"
End Sub


Private Sub Attempt16txt_Click()
Statuslbl = "Crack Attempt 16"
AnswerCrack.DialBox = AnswerCrack.Attempt16txt
AnswerCrack.Attempt16txt.Enabled = False
AttemptNumber = 16
Command1_Click
Command1.Caption = "Play Attempt 16"
End Sub

Private Sub Attempt17txt_Click()
Statuslbl = "Crack Attempt 17"
AnswerCrack.DialBox = AnswerCrack.Attempt17txt
AnswerCrack.Attempt17txt.Enabled = False
AttemptNumber = 17
Command1_Click
Command1.Caption = "Play Attempt 17"
End Sub

Private Sub Attempt18txt_Click()
Statuslbl = "Crack Attempt 18"
AnswerCrack.DialBox = AnswerCrack.Attempt18txt
AnswerCrack.Attempt18txt.Enabled = False
AttemptNumber = 18
Command1_Click
Command1.Caption = "Play Attempt 18"
End Sub

Private Sub Attempt19txt_Click()
Statuslbl = "Crack Attempt 19"
AnswerCrack.DialBox = AnswerCrack.Attempt19txt
AnswerCrack.Attempt19txt.Enabled = False
AttemptNumber = 19
Command1_Click
Command1.Caption = "Play Attempt 19"
End Sub

Private Sub Attempt20txt_Click()
Statuslbl = "Crack Attempt 20"
AnswerCrack.DialBox = AnswerCrack.Attempt20txt
AnswerCrack.Attempt20txt.Enabled = False
AttemptNumber = 20
Command1_Click
Command1.Caption = "Play Attempt 20"
End Sub

Private Sub Combotxt_Change()
'On Error GoTo SkipThis:
'Itemz = Combotxt
'Check = Len(Itemz)
'If Check = 1 Or Check = 2 Then Exit Sub
'DoEvents
'List2.RemoveItem Itemz
'DoEvents
'SkipThis:
'Resume 1
'1 :
End Sub

Private Sub Combotxt_Click()
Statuslbl = "Crack Attempt 18"
AnswerCrack.DialBox = Combotxt
End Sub

Private Sub Attempt1txt_Click()
Statuslbl = "Crack Attempt 1"
AnswerCrack.DialBox = AnswerCrack.Attempt1txt
AnswerCrack.Attempt1txt.Enabled = False
AttemptNumber = 1
Command1_Click
Command1.Caption = "Play Attempt 1"
End Sub

Private Sub Attempt2txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt2txt
Statuslbl = "Crack Attempt 2"
AnswerCrack.Attempt2txt.Enabled = False
AttemptNumber = 2
Command1_Click
Command1.Caption = "Play Attempt 2"
End Sub

Private Sub Attempt3txt_Click()
Statuslbl = "Crack Attempt 3"
AnswerCrack.DialBox = AnswerCrack.Attempt3txt
AnswerCrack.Attempt3txt.Enabled = False
AttemptNumber = 3
Command1_Click
Command1.Caption = "Play Attempt 3"
End Sub

Private Sub Attempt4txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt4txt
Statuslbl = "Crack Attempt 4"
AnswerCrack.Attempt4txt.Enabled = False
AttemptNumber = 4
Command1_Click
Command1.Caption = "Play Attempt 4"
End Sub

Private Sub Attempt5txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt5txt
Statuslbl = "Crack Attempt 5"
AnswerCrack.Attempt5txt.Enabled = False
AttemptNumber = 5
Command1_Click
Command1.Caption = "Play Attempt 5"
End Sub

Private Sub Attempt6txt_Click()
AnswerCrack.DialBox = AnswerCrack.Attempt6txt
Statuslbl = "Crack Attempt 6"
AnswerCrack.Attempt6txt.Enabled = False
AttemptNumber = 6
Command1_Click
Command1.Caption = "Play Attempt 6"
End Sub

Private Sub thirteensec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
Statuslbl = "13 Sections Approx - 10 Second Intervals"
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "370380390300300410420430440450460470480490400500510520530540550560570580590506006106206306"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "4065066067068069060700710720730740750760770780790708008108208308408508608708808908090091"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "092093094095096097098099099111113131132133211212221222231232233311312313321322323331332333"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "411412413414421422423424431432433434441442443444511512513514515521522523524525531532533534"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "535541542543544545551552553554555611612613614615616621622623624625626631632633634635636641"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "642643644645646651652653654655656661662663664665666711712713714715716717721722723724725726"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "727731732733734735736737741742743744745746747751752753754755756757761762763764765766767771"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "772773774775776777811812813814815816817818821822823824825826827828831832833834835836837838"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "841842843844845846847848851852853854855856857858861862863864865866867868871872873874875876"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "877878881882883884885886887888911912913914915916917918919921922923924925926927928929931932"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "933934935936937938939941942943944945946947948949951952953954955956957958959961962963964965"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt13txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)

                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a      '& Box13a & Box25b & Box13b & Box25a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt13txt = AnswerCrack.Attempt13txt & Extras



End Sub

Private Sub ThreeDigit_Click()
AnswerCrack.Combotxt = "3 Digit Combination"
AnswerCrack.section.Enabled = True

List1.Clear
If threedigit.Checked = True Then
threedigit.Checked = False

AnswerCrack.Attempt1txt.Visible = False
AnswerCrack.Attempt2txt.Visible = False
AnswerCrack.Attempt3txt.Visible = False
AnswerCrack.Attempt4txt.Visible = False
AnswerCrack.Attempt5txt.Visible = False
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
Exit Sub
End If
If threedigit.Checked = False Then
threedigit.Checked = True
twodigit.Checked = False
AnswerCrack.DialBox = ""
'AnswerCrack.Attempt1txt.Visible = True
'AnswerCrack.Attempt2txt.Visible = True
'AnswerCrack.Attempt3txt.Visible = True
'AnswerCrack.Attempt4txt.Visible = True
'AnswerCrack.Attempt5txt.Visible = True
'AnswerCrack.Attempt6txt.Visible = True
'AnswerCrack.Attempt7txt.Visible = True
'AnswerCrack.Attempt8txt.Visible = True
'AnswerCrack.Attempt9txt.Visible = True
'AnswerCrack.Attempt10txt.Visible = True
'AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.DialBox = "0010020030040050060070080090100110120130140150160170180190200210220230240250260270280290300310320330340350360370380390400410420430440450460470480490500510520530540550560570580590600610620630640650660670680690700710720730740750760770780790800810820830840850860870880890900910920930940950960970980991111131311321332112122212222312322333113123133213223233313323334114124134144214224234244314324"
AnswerCrack.DialBox = AnswerCrack.DialBox + "33434441442443444511512513514515521522523524525531532533534535541542543544545551552553554555611612613614615616621622623624625626631632633634635636641642643644645646651652653654655656661662663664665666711712713714715716717721722723724725726727731732733734735736737741742743744745746747751752753754755756757761762763764765766767771772773774775776777811812813814815816817818821822823824825826827"
AnswerCrack.DialBox = AnswerCrack.DialBox + "828831832833834835836837838841842843844845846847848851852853854855856857858861862863864865866867868871872873874875876877878881882883884885886887888911912913914915916917918919921922923924925926927928929931932933934935936937938939941942943944945946947948949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"

Exit Sub
End If

End Sub

Private Sub ThreeSec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call pause(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call pause(2): Statuslbl = "": Exit Sub
Statuslbl = "3 Sections Approx - 45 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = False
AnswerCrack.Attempt5txt.Visible = False
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
AnswerCrack.Attempt12txt.Visible = False
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False

'000, 901 fragmented missed before also
'901 fixed in first strand
'020 now missing
'902 now missing

AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250260270280290230031032033034035036037038039030041042043044045046047048049040050051052053054055056057058059050600610620630640650660670680690607007107207307407507607707807907080081082083084085086087088089080900910920930940950960970980990991111131311321332112122212222312322333113123133213223233313323334114124134144214224234244314324334344414424434445"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "1151251351451552152252352452553153253353453554154254354454555155255355455561161261361461561662162262362462562663163263363463563664164264364464564665165265365465565666166266366466566671171271371471571671772172272372472572672773173273373473573673774174274374474574674775175275375475575675776176276376476576676777177277377477577677781181281381481581681781882182282382482582682782883183"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "2833834835836837838841842843844845846847848851852853854855856857858861862863864865866867868871872873874875876877878881882883884885886887888911912913914915916917918919921922923924925926927928929931932933934935936937938939941942943944945946947948949951952953954955956957958959961962963964965966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999243433278782"
AnswerCrack.Attempt3txt.Enabled = True

AnswerCrack.Attempt4txt.Visible = False
AnswerCrack.Attempt5txt.Visible = False
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr

'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)  '12 Box 3 First 2 Chr
Box19b = Left(AnswerCrack.Attempt8txt, 2)  '12 Box 3 First 2 Chr
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
                      
        '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a        '& Box3a   '& Box15b & Box3b & Box15a '& Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box15b & Box7b & Box15a & Box8a & Box20b & Box8b & Box20a & Box8a & Box19b & Box8b & Box19a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a
Text10 = Extras

'Add On Extras to Last Box
AnswerCrack.Attempt3txt = AnswerCrack.Attempt3txt & Extras

End Sub



Private Sub Timer1_Timer()
'This timer keeps the list's locked when scrolling
Call ListLock(List1, List2, True)
End Sub

Private Sub twelvesec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "12 Sections Approx - 11 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = False
AnswerCrack.Attempt14txt.Visible = False
AnswerCrack.Attempt15txt.Visible = False
AnswerCrack.Attempt16txt.Visible = False
AnswerCrack.Attempt17txt.Visible = False
AnswerCrack.Attempt18txt.Visible = False
AnswerCrack.Attempt19txt.Visible = False
AnswerCrack.Attempt20txt.Visible = False
                         '  000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390300410
AnswerCrack.Attempt1txt = "000100110120130140150160170180190102012002102202302402502602702802902300310320330340350360370380390"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "30030041042043044045046047048049040050051052053054055056057058059050600610620630640650"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "66067068069060700710720730740750760770780790708008108208308408508608708808908090091092093094095096"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "09709809909911111313113213321121222122223123223331131231332132232333133233341141241341442142242342443143243"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "34344414424434445115125135145155215225235245255315325335345355415425435445455515525535545556116126"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "13614615616621622623624625626631632633634635636641642643644645646651652653654655656661662663664665"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "66671171271371471571671772172272372472572672773173273373473573673774174274374474574674775175275375"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "47557567577617627637647657667677717727737747757767778118128138148158168178188218228238248258268278"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "28831832833834835836837838841842843844845846847848851852853854855856857858861862863864865866867868"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "87187287387487587687787888188288388488588688788891191291391491591691791891992192292392492592692792"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "89299319329339349359369379389399419429439449459469479489499519529539549559569579589599619629639649"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "65966967968969971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt12txt.Enabled = True


Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
Box12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
Box12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)

                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a '& Box12a & Box24b & Box12b & Box24a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt12txt = AnswerCrack.Attempt12txt & Extras




End Sub

Private Sub twentysec_Click()
AnswerCrack.List1.Clear
If twodigit.Checked = False And threedigit.Checked = False Then Statuslbl = "Please Select Amount Of Digits In Code": Call Timeout(2): Statuslbl = "": Exit Sub
If twodigit.Checked = True Then Statuslbl = "Two Digit Codes Do Not Need To Be Segmented": Call Timeout(2): Statuslbl = "": Exit Sub
Statuslbl = "20 Sections Approx - 6 Second Intervals"
AnswerCrack.Attempt1txt = ""
AnswerCrack.Attempt2txt = ""
AnswerCrack.Attempt3txt = ""
AnswerCrack.Attempt4txt = ""
AnswerCrack.Attempt5txt = ""
AnswerCrack.Attempt6txt = ""
AnswerCrack.Attempt7txt = ""
AnswerCrack.Attempt8txt = ""
AnswerCrack.Attempt9txt = ""
AnswerCrack.Attempt10txt = ""
AnswerCrack.Attempt11txt = ""
AnswerCrack.Attempt12txt = ""
AnswerCrack.Attempt13txt = ""
AnswerCrack.Attempt14txt = ""
AnswerCrack.Attempt15txt = ""
AnswerCrack.Attempt16txt = ""
AnswerCrack.Attempt17txt = ""
AnswerCrack.Attempt18txt = ""
AnswerCrack.Attempt19txt = ""
AnswerCrack.Attempt20txt = ""
AnswerCrack.Attempt1txt.Visible = True
AnswerCrack.Attempt2txt.Visible = True
AnswerCrack.Attempt3txt.Visible = True
AnswerCrack.Attempt4txt.Visible = True
AnswerCrack.Attempt5txt.Visible = True
AnswerCrack.Attempt6txt.Visible = True
AnswerCrack.Attempt7txt.Visible = True
AnswerCrack.Attempt8txt.Visible = True
AnswerCrack.Attempt9txt.Visible = True
AnswerCrack.Attempt10txt.Visible = True
AnswerCrack.Attempt11txt.Visible = True
AnswerCrack.Attempt12txt.Visible = True
AnswerCrack.Attempt13txt.Visible = True
AnswerCrack.Attempt14txt.Visible = True
AnswerCrack.Attempt15txt.Visible = True
AnswerCrack.Attempt16txt.Visible = True
AnswerCrack.Attempt17txt.Visible = True
AnswerCrack.Attempt18txt.Visible = True
AnswerCrack.Attempt19txt.Visible = True
AnswerCrack.Attempt20txt.Visible = True



AnswerCrack.Attempt1txt = "00010011012013014015016017018019010201200210220230240250"
AnswerCrack.Attempt1txt.Enabled = True
AnswerCrack.Attempt2txt = "26027028029023003103203303403503603703803903003004104204"
AnswerCrack.Attempt2txt.Enabled = True
AnswerCrack.Attempt3txt = "30440450460470480490400500510520530540550560570580590506"
AnswerCrack.Attempt3txt.Enabled = True
AnswerCrack.Attempt4txt = "00610620630640650660670680690607007107207307407507607707"
AnswerCrack.Attempt4txt.Enabled = True
AnswerCrack.Attempt5txt = "80790708008108208308408508608708808908090091092093094095"
AnswerCrack.Attempt5txt.Enabled = True
AnswerCrack.Attempt6txt = "09609709809909911111313113213321121222122223123223331131"
AnswerCrack.Attempt6txt.Enabled = True
AnswerCrack.Attempt7txt = "23133213223233313323334114124134144214224234244314324334"
AnswerCrack.Attempt7txt.Enabled = True
AnswerCrack.Attempt8txt = "34441442443444511512513514515521522523524525531532533534"
AnswerCrack.Attempt8txt.Enabled = True
AnswerCrack.Attempt9txt = "53554154254354454555155255355455561161261361461561662162"
AnswerCrack.Attempt9txt.Enabled = True
AnswerCrack.Attempt10txt = "2623624625626631632633634635636641642643644645646651652"
AnswerCrack.Attempt10txt.Enabled = True
AnswerCrack.Attempt11txt = "653654655656661662663664665666711712713714715716717721722723724"
AnswerCrack.Attempt11txt.Enabled = True
AnswerCrack.Attempt12txt = "72572672773173273373473573673774174274374474574674775175275375"
AnswerCrack.Attempt12txt.Enabled = True
AnswerCrack.Attempt13txt = "4755756757761762763764765766767771772773774775776777811"
AnswerCrack.Attempt13txt.Enabled = True
AnswerCrack.Attempt14txt = "81281381481581681781882182282382482582682782883183283383483583"
AnswerCrack.Attempt14txt.Enabled = True
AnswerCrack.Attempt15txt = "68378388418428438448458468478488518528538548558568578588618"
AnswerCrack.Attempt15txt.Enabled = True
AnswerCrack.Attempt16txt = "6286386486586686786887187287387487587687787888188288388488588688788891191291"
AnswerCrack.Attempt16txt.Enabled = True
AnswerCrack.Attempt17txt = "39149159169179189199219229239249259269279289299319329339349359369379"
AnswerCrack.Attempt17txt.Enabled = True
AnswerCrack.Attempt18txt = "3893994194294394494594694794894995195295395495595695795895996196296396496596696796896"
AnswerCrack.Attempt18txt.Enabled = True
AnswerCrack.Attempt19txt = "9971972973974975976977978979981982983984985986987988989991992993994995996997998999"
AnswerCrack.Attempt19txt.Enabled = True
AnswerCrack.Attempt20txt = ""
AnswerCrack.Attempt20txt.Enabled = True

Box1a = Right(AnswerCrack.Attempt1txt, 1) ' 4Box 1 Last 1 Chr
Box1b = Right(AnswerCrack.Attempt1txt, 2) ' 4Box 1 Last 2 Chr
Box2a = Right(AnswerCrack.Attempt2txt, 1) ' 5Box 2 Last 1 Chr
Box2b = Right(AnswerCrack.Attempt2txt, 2) ' 5Box 2 Last 2 Chr
Box3a = Right(AnswerCrack.Attempt3txt, 1) ' 6Box 3 Last 1 Chr
Box3b = Right(AnswerCrack.Attempt3txt, 2) ' 6Box 3 Last 2 Chr
Box4a = Right(AnswerCrack.Attempt4txt, 1) ' 7Box 4 Last 1 Chr
Box4b = Right(AnswerCrack.Attempt4txt, 2) ' 7Box 4 Last 2 Chr
Box5a = Right(AnswerCrack.Attempt5txt, 1) ' 8Box 5 Last 1 Chr
Box5b = Right(AnswerCrack.Attempt5txt, 2) ' 8Box 5 Last 2 Chr
Box6a = Right(AnswerCrack.Attempt6txt, 1) ' 9Box 6 Last 1 Chr
Box6b = Right(AnswerCrack.Attempt6txt, 2) ' 9Box 6 Last 2 Chr
Box7a = Right(AnswerCrack.Attempt7txt, 1)  ' 11Box 3 First 2 Chr
Box7b = Right(AnswerCrack.Attempt7txt, 2)  ' 11Box 3 First 2 Chr
Box8a = Right(AnswerCrack.Attempt8txt, 1)
Box8b = Right(AnswerCrack.Attempt8txt, 2)
Box9a = Right(AnswerCrack.Attempt9txt, 1)  ' 13Box 3 First 2 Chr
Box9b = Right(AnswerCrack.Attempt9txt, 2)  ' 13Box 3 First 2 Chr
Box10a = Right(AnswerCrack.Attempt10txt, 1)  ' 14Box 3 First 2 Chr
Box10b = Right(AnswerCrack.Attempt10txt, 2)  ' 14Box 3 First 2 Chr
Box11a = Right(AnswerCrack.Attempt11txt, 1)  ' 15Box 3 First 2 Chr
Box11b = Right(AnswerCrack.Attempt11txt, 2)  ' 15Box 3 First 2 Chr
RightBox12a = Right(AnswerCrack.Attempt12txt, 1)  ' 15Box 3 First 2 Chr
RightBox12b = Right(AnswerCrack.Attempt12txt, 2)  ' 15Box 3 First 2 Chr
RightBox13a = Right(AnswerCrack.Attempt13txt, 1)  ' 15Box 3 First 2 Chr
RightBox13b = Right(AnswerCrack.Attempt13txt, 2)  ' 15Box 3 First 2 Chr
RightBox14a = Right(AnswerCrack.Attempt14txt, 1)  ' 15Box 3 First 2 Chr
RightBox14b = Right(AnswerCrack.Attempt14txt, 2)  ' 15Box 3 First 2 Chr
RightBox15a = Right(AnswerCrack.Attempt15txt, 1)  ' 15Box 3 First 2 Chr
RightBox15b = Right(AnswerCrack.Attempt15txt, 2)  ' 15Box 3 First 2 Chr
RightBox16a = Right(AnswerCrack.Attempt16txt, 1)  ' 15Box 3 First 2 Chr
RightBox16b = Right(AnswerCrack.Attempt16txt, 2)  ' 15Box 3 First 2 Chr
RightBox17a = Right(AnswerCrack.Attempt17txt, 1)  ' 15Box 3 First 2 Chr
RightBox17b = Right(AnswerCrack.Attempt17txt, 2)  ' 15Box 3 First 2 Chr
RightBox18a = Right(AnswerCrack.Attempt18txt, 1)  ' 15Box 3 First 2 Chr
RightBox18b = Right(AnswerCrack.Attempt18txt, 2)  ' 15Box 3 First 2 Chr
RightBox19a = Right(AnswerCrack.Attempt19txt, 1)  ' 15Box 3 First 2 Chr
RightBox19b = Right(AnswerCrack.Attempt19txt, 2)  ' 15Box 3 First 2 Chr


'-----------------------
Box12a = Left(AnswerCrack.Attempt1txt, 1)  ' 4Box 1 First Chr
Box12b = Left(AnswerCrack.Attempt1txt, 2)  ' 4Box 1 First 2 Chr
Box13a = Left(AnswerCrack.Attempt2txt, 1)  ' 5Box 2 First Chr
Box13b = Left(AnswerCrack.Attempt2txt, 2)  ' 5Box 2 First 2 Chr
Box14a = Left(AnswerCrack.Attempt3txt, 1)  ' 6Box 3 First Chr
Box14b = Left(AnswerCrack.Attempt3txt, 2)  ' 6Box 3 First 2 Chr
Box15a = Left(AnswerCrack.Attempt4txt, 1)  ' 7Box 3 First 2 Chr
Box15b = Left(AnswerCrack.Attempt4txt, 2)  ' 7Box 3 First 2 Chr
Box16a = Left(AnswerCrack.Attempt5txt, 1)  ' 8Box 3 First 2 Chr
Box16b = Left(AnswerCrack.Attempt5txt, 2)  ' 8Box 3 First 2 Chr
Box17a = Left(AnswerCrack.Attempt6txt, 1)  ' 9Box 3 First 2 Chr
Box17b = Left(AnswerCrack.Attempt6txt, 2)  ' 9Box 3 First 2 Chr
Box18a = Left(AnswerCrack.Attempt7txt, 1)  '11 Box 3 First 2 Chr
Box18b = Left(AnswerCrack.Attempt7txt, 2)  '11 Box 3 First 2 Chr
Box19a = Left(AnswerCrack.Attempt8txt, 1)
Box19b = Left(AnswerCrack.Attempt8txt, 2)
Box20a = Left(AnswerCrack.Attempt9txt, 1)  '13 Box 3 First 2 Chr
Box20b = Left(AnswerCrack.Attempt9txt, 2)  '13 Box 3 First 2 Chr
Box21a = Left(AnswerCrack.Attempt10txt, 1)  '14 Box 3 First 2 Chr
Box21b = Left(AnswerCrack.Attempt10txt, 2)  '14 Box 3 First 2 Chr
Box22a = Left(AnswerCrack.Attempt11txt, 1)  '15 Box 3 First 2 Chr
Box22b = Left(AnswerCrack.Attempt11txt, 2)  '15 Box 3 First 2 Chr
Box23a = Left(AnswerCrack.Attempt12txt, 1)
Box23b = Left(AnswerCrack.Attempt12txt, 2)
Box24a = Left(AnswerCrack.Attempt13txt, 1)
Box24b = Left(AnswerCrack.Attempt13txt, 2)
Box25a = Left(AnswerCrack.Attempt14txt, 1)
Box25b = Left(AnswerCrack.Attempt14txt, 2)
Box26a = Left(AnswerCrack.Attempt15txt, 1)
Box26b = Left(AnswerCrack.Attempt15txt, 2)
Box27a = Left(AnswerCrack.Attempt16txt, 1)
Box27b = Left(AnswerCrack.Attempt16txt, 2)
Box28a = Left(AnswerCrack.Attempt17txt, 1)
Box28b = Left(AnswerCrack.Attempt17txt, 2)
Box29a = Left(AnswerCrack.Attempt18txt, 1)
Box29b = Left(AnswerCrack.Attempt18txt, 2)
Box30a = Left(AnswerCrack.Attempt19txt, 1)
Box30b = Left(AnswerCrack.Attempt19txt, 2)
Box31a = Left(AnswerCrack.Attempt20txt, 1)
Box31b = Left(AnswerCrack.Attempt20txt, 2)


                              '1chr  2First 4         1chr  4     First          5       First         5                     6                       6                        7                     7                      8                      8                        9                9                      11                    11                       12                   12                          13                        13                  14                     14                     15                         15                    16                        16                       17                        17                    18                 18
Extras = Box1a & Box13b & Box1b & Box13a & Box2a & Box14b & Box2b & Box14a & Box3a & Box15b & Box3b & Box15a & Box4a & Box16b & Box4b & Box16a & Box5a & Box17b & Box5b & Box17a & Box6a & Box18b & Box6b & Box18a & Box7a & Box19b & Box7b & Box19a & Box8a & Box20b & Box8b & Box20a & Box9a & Box21b & Box9b & Box21a & Box10a & Box22b & Box10b & Box22a & Box11a & Box23b & Box11b & Box23a & RightBox12a & Box24b & RightBox12b & Box24a & RightBox13a & Box25b & RightBox13b & Box25a & RightBox14a & Box26b & RightBox14b & Box26a & RightBox15a & Box27b & RightBox15b & Box27a & RightBox16a & Box28b & RightBox16b & Box28a & RightBox17a & Box29b & RightBox17b & Box29a & RightBox18a & Box30b & RightBox18b & Box30a & RightBox19a & Box31b & RightBox19b & Box31a
Text10 = Extras
'Add On Extras to Last Box
AnswerCrack.Attempt20txt = AnswerCrack.Attempt20txt & Extras


End Sub

Private Sub TwoDigit_Click()
AnswerCrack.Combotxt = "2 Digit Combination"
AnswerCrack.section.Enabled = False
List1.Clear
If twodigit.Checked = False Then
twodigit.Checked = True
threedigit.Checked = False
AnswerCrack.DialBox = "001122334455667788991357902468036925814715937049483827261605173950628408529630074197531864209876543210"
AnswerCrack.Attempt1txt.Visible = False
AnswerCrack.Attempt2txt.Visible = False
AnswerCrack.Attempt3txt.Visible = False
AnswerCrack.Attempt4txt.Visible = False
AnswerCrack.Attempt5txt.Visible = False
AnswerCrack.Attempt6txt.Visible = False
AnswerCrack.Attempt7txt.Visible = False
AnswerCrack.Attempt8txt.Visible = False
AnswerCrack.Attempt9txt.Visible = False
AnswerCrack.Attempt10txt.Visible = False
AnswerCrack.Attempt11txt.Visible = False
Exit Sub
End If

If twodigit.Checked = True Then
twodigit.Checked = False
AnswerCrack.DialBox = ""
'AnswerCrack.Attempt1txt.Visible = True
'AnswerCrack.Attempt2txt.Visible = True
'AnswerCrack.Attempt3txt.Visible = True
'AnswerCrack.Attempt4txt.Visible = True
'AnswerCrack.Attempt5txt.Visible = True
'AnswerCrack.Attempt6txt.Visible = True
'AnswerCrack.Attempt7txt.Visible = True
'AnswerCrack.Attempt8txt.Visible = True
'AnswerCrack.Attempt9txt.Visible = True
'AnswerCrack.Attempt10txt.Visible = True
'AnswerCrack.Attempt11txt.Visible = True
Exit Sub
End If

End Sub

