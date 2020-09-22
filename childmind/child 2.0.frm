VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   3645
   Icon            =   "child 2.0.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2760
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   855
      Left            =   1440
      Picture         =   "child 2.0.frx":0442
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":213C
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1200
      Picture         =   "child 2.0.frx":257E
      Tag             =   "key"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":29C0
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1920
      Picture         =   "child 2.0.frx":2E02
      Tag             =   "heart"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":3244
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   480
      Picture         =   "child 2.0.frx":3686
      Tag             =   "bell"
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":3AC8
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1200
      Picture         =   "child 2.0.frx":3F0A
      Tag             =   "smile"
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":434C
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   2640
      Picture         =   "child 2.0.frx":4796
      Tag             =   "mic"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":4BE0
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   1920
      Picture         =   "child 2.0.frx":5022
      Tag             =   "house"
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":5464
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   2640
      Picture         =   "child 2.0.frx":58A6
      Tag             =   "eye"
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragIcon        =   "child 2.0.frx":5CE8
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   480
      Picture         =   "child 2.0.frx":5FF2
      Tag             =   "camera"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 To   Play New         Game."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Menu Game 
      Caption         =   "&Game"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seconds As Integer

Private Sub exit_Click()
End
End Sub

Private Sub Image9_DragDrop(Source As Control, X As Single, Y As Single)

If Source.Tag = "camera" Then
Label2.Caption = 10
Image1.Visible = False
Label1.Caption = "Eye"
Image2.Enabled = True

ElseIf Source.Tag = "eye" Then
Label2.Caption = Label2.Caption + 10
Image2.Visible = False
Label1.Caption = "House"
Image3.Enabled = True

ElseIf Source.Tag = "house" Then
Label2.Caption = Label2.Caption + 10
Image3.Visible = False
Label1.Caption = "Mike"
Image4.Enabled = True

ElseIf Source.Tag = "mic" Then
Label2.Caption = Label2.Caption + 10
Image4.Visible = False
Label1.Caption = "Smile"
Image5.Enabled = True

ElseIf Source.Tag = "smile" Then
Label2.Caption = Label2.Caption + 10
Image5.Visible = False
Label1.Caption = "Bell"
Image6.Enabled = True


ElseIf Source.Tag = "bell" Then
Label2.Caption = Label2.Caption + 10
Image6.Visible = False
Label1.Caption = "Heart"
Image7.Enabled = True


ElseIf Source.Tag = "heart" Then
Label2.Caption = Label2.Caption + 10
Image7.Visible = False
Label1.Caption = "Key"
Image8.Enabled = True

ElseIf Source.Tag = "key" Then
Label2.Caption = Label2.Caption + 10
Image8.Visible = False
Label1.Caption = "Press F2 To Start a new game."
Timer1.Enabled = False

End If

End Sub

Private Sub new_Click()
Label1.Caption = "Camera"
Label2.Caption = " "
Label4.Caption = "Score"
Label2.Caption = "0"
Label5.Caption = "Time left"
Label6.Caption = "60"
Label3.Visible = False
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Image5.Visible = True
Image6.Visible = True
Image7.Visible = True
Image8.Visible = True
Image9.Visible = True
Image1.Enabled = True
Image2.Enabled = False
Image3.Enabled = False
Image4.Enabled = False
Image5.Enabled = False
Image6.Enabled = False
Image7.Enabled = False
Image8.Enabled = False
Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()

If seconds = 0 Then
  seconds = 60
Else
 seconds = seconds - 1
Label6.Caption = seconds

End If
If Label6.Caption = "0" Then
Label1.Caption = "Time Is Up! Press F2."
Timer1.Enabled = False
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
Image9.Visible = False

End If
End Sub


