VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPaint 
   Caption         =   "Picture Painter"
   ClientHeight    =   5130
   ClientLeft      =   345
   ClientTop       =   1050
   ClientWidth     =   8850
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8850
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cool!"
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   4680
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opaque"
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
      Begin VB.OptionButton optopaq 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optopaq 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdObject 
      Caption         =   "Rectangle"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdObject 
      Caption         =   "Circle"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdObject 
      Caption         =   "Line"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdObject 
      Caption         =   "Dot"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog diag 
      Left            =   960
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicFPreview 
      Height          =   255
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox PicBPreview 
      Height          =   255
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtRad 
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "300"
      Top             =   4680
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Border"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
      Begin VB.OptionButton optBorder 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optBorder 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   1440
      ScaleHeight     =   4995
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Paint Type"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Border/line Color"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Filll Color"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Radius of circles"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileClear 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Description: A paint program with limited features
' November 23, 2000
' Programmed by Paul Bergman
' E-mail me at: pbergman001@hotmail.com
' You may feel free to add/edit the code. For example add a save Menu

Dim Rad As Long
Dim X1 As Long, X2 As Long
Dim Y1 As Long, Y2 As Long
Dim Types As String
Dim Ctr As Long
Dim SavePic
Dim Xpos As Long
Dim Ypos As Long

Private Sub cmdObject_Click(Index As Integer)
  ' Assigns a value to Types (Decides what type of image will be displayed)
  Select Case Index
    Case 0
      Types = "dot"
    Case 1
      Types = "line"
    Case 2
      Types = "circle"
    Case 3
      Types = "box"
    Case Else
      MsgBox Index
  End Select
  Ctr = 1
End Sub

Private Sub Command1_Click()
  Timer1.Enabled = Not Timer1.Enabled  ' On/Off switch for the Timer
  Timer2.Enabled = False
End Sub

Private Sub Form_Load()
  Randomize Timer
  PicFPreview.BackColor = vbBlue
  PicBPreview.BackColor = vbBlack
  ChangeColor
  optBorder(0) = True
  optopaq(0) = True
End Sub

Private Sub mnuFileClear_Click()
  Picture1.Cls
End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub optBorder_Click(Index As Integer)
  If optBorder(0) = True Then
    Picture1.ForeColor = PicBPreview.BackColor
  ElseIf optBorder(1) = True Then
    Picture1.ForeColor = Picture1.FillColor
  End If
End Sub

Private Sub optopaq_Click(Index As Integer)
  Picture1.FillStyle = Index
End Sub

Private Sub PicBPreview_Click()
  Me.diag.ShowColor
  PicBPreview.BackColor = Me.diag.Color
  ChangeColor
End Sub

Private Sub PicFPreview_Click()
  Me.diag.ShowColor
  PicFPreview.BackColor = Me.diag.Color
  ChangeColor
End Sub

Private Sub ChangeColor()
  Picture1.ForeColor = PicBPreview.BackColor
  Picture1.FillColor = PicFPreview.BackColor
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Types
    Case "dot"
      Timer3.Enabled = True
    Case "line"
      If Ctr = 1 Then     ' Make sure that there are a first set of coordinates
        X2 = X1
        Y2 = Y1
        Ctr = Ctr + 1
      ElseIf Ctr = 2 Then  ' Gathers second set of coordinates to complete the line
        Picture1.Line (X2, Y2)-(X1, Y1)
        Ctr = 1
      End If
    Case "circle"
      Timer3.Enabled = True
    Case "box"
      If Ctr = 1 Then   ' Make sure that there are a first set of coordinates
        X2 = X1
        Y2 = Y1
        Ctr = Ctr + 1
      ElseIf Ctr = 2 Then ' Gathers second set of coordinates to complete the Box
        Picture1.Line (X2, Y2)-(X1, Y1), , B
        Ctr = 1           ' resets the counter to gather 1st set of coordinates
      End If
    Case Else
      MsgBox "Please choose a Paint Type!", vbCritical, "ERROR"
  End Select
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  X1 = X    'Variables used to always keep X,Y positions
  Y1 = Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer3.Enabled = False
End Sub

Private Sub Timer1_Timer()

  If Xpos >= Picture1.Width Or Ypos >= Picture1.Height Then
    Timer2.Enabled = True    ' Used as a switch
    Timer1.Enabled = False
    Timer2_Timer
  End If
  Xpos = Xpos + (Rnd * 70)    ' creates a random value for movement
  Ypos = Ypos + (Rnd * 70)
  
  Picture1.Circle (Xpos, Ypos), 300 'Creates a circle
  
End Sub

Private Sub Timer2_Timer()
  If Xpos <= 0 Or Ypos <= 0 Then
    Timer2.Enabled = False    ' Part Two of the switch
    Timer1.Enabled = True
    Timer1_Timer
  End If
  Xpos = Xpos - (Rnd * 70)
  Ypos = Ypos - (Rnd * 70)
  Picture1.Circle (Xpos, Ypos), 300
End Sub

Private Sub Timer3_Timer()
  Picture1.PSet (X1, Y1)
  
  Select Case Types
    Case "dot"
      Picture1.PSet (X1, Y1)  'Will draw points on the picture for as long as the mouseDown is in effect
    Case "circle"
      optBorder_Click (Index)
      If txtRad = "" Then
        MsgBox ("Please enter a radius")
        Exit Sub
      End If
      Rad = Val(txtRad)
      Picture1.Circle (X1, Y1), Rad   'Will draw Circles
  End Select
End Sub
