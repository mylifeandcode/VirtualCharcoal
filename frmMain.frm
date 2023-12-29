VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Charcoal"
   ClientHeight    =   3165
   ClientLeft      =   1320
   ClientTop       =   1200
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3705
   Begin VB.Timer tmrNeeds 
      Interval        =   60000
      Left            =   1680
      Top             =   3000
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   500
      Left            =   960
      Top             =   3000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdMedicine 
      Height          =   615
      Left            =   2760
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Give Medicine"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdEmptyLitter 
      Height          =   615
      Left            =   2760
      Picture         =   "frmMain.frx":0CB2
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Empty Litter"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdPlay 
      Height          =   615
      Left            =   2760
      Picture         =   "frmMain.frx":1964
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Play"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdFeed 
      Height          =   615
      Left            =   2760
      Picture         =   "frmMain.frx":2616
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Feed"
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox picCharcoal 
      Height          =   1860
      Left            =   240
      Picture         =   "frmMain.frx":32C8
      ScaleHeight     =   1800
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Virtual Charcoal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Charcoal As Cat
  Dim sTimeOfDeath As String

Private Sub cmdEmptyLitter_Click()
  If Charcoal.bNeedsLitter = True Then
    Charcoal.bNeedsLitter = False
    Charcoal.iFrame = 0
  End If
End Sub

Private Sub cmdExit_Click()
  Dim counter
  For counter = 1 To 300
    Beep
  Next
  End
End Sub

Private Sub cmdFeed_Click()
  If Charcoal.bNeedsFood = True Then
    Charcoal.bNeedsFood = False
    DetermineFrame
  End If
End Sub

Private Sub cmdMedicine_Click()
  If Charcoal.bIsSick = True Then
    Charcoal.bIsSick = False
    DetermineFrame
  End If
End Sub

Private Sub cmdPlay_Click()
  If Charcoal.bNeedsToPlay = True Then
    Charcoal.bNeedsToPlay = False
    DetermineFrame
  End If
End Sub

Private Sub Form_Load()
  'Initialize Charcoal's Values
  With Charcoal
    .bAlive = True
    .bNeedsFood = False
    .bNeedsToPlay = False
    .bNeedsLitter = False
    .bIsSick = False
    .iHappiness = 75
    .iHealth = 100
    .iFrame = 0
  End With
  
  Randomize
End Sub

Private Sub tmrAnimation_Timer()
  Dim i As Integer
  Static iLastFrame As Integer
  Select Case Charcoal.iFrame
    Case 0
        'OK Frame 1
        Charcoal.iFrame = 1
        iLastFrame = 0
    Case 1
        'OK Frame 2
        If iLastFrame = 2 Then
            Charcoal.iFrame = 0
        Else
            Charcoal.iFrame = 2
        End If
        iLastFrame = 1
    Case 2
        'OK Frame 3
        Charcoal.iFrame = 1
        iLastFrame = 2
    Case 3
        'Needs Food Frame 1
        Charcoal.iFrame = 4
        iLastFrame = 3
    Case 4
        'Needs Food Frame 2
        If iLastFrame = 3 Then
            Charcoal.iFrame = 5
        Else
            Charcoal.iFrame = 3
        End If
        iLastFrame = 4
    Case 5
        'Needs Food Frame 3
        Charcoal.iFrame = 4
        iLastFrame = 5
    Case 6
        'Needs Playtime Frame 1
        Charcoal.iFrame = 7
        iLastFrame = 6
    Case 7
        'Needs Playtime Frame 2
        If iLastFrame = 6 Then
            Charcoal.iFrame = 8
        Else
            Charcoal.iFrame = 6
        End If
        iLastFrame = 7
    Case 8
        'Needs Playtime Frame 3
        Charcoal.iFrame = 7
        iLastFrame = 8
    Case 9
        'Needs New Litter Frame 1
        Charcoal.iFrame = 10
        iLastFrame = 9
    Case 10
        'Needs New Litter Frame 2
        If iLastFrame = 9 Then
            Charcoal.iFrame = 11
        Else
            Charcoal.iFrame = 9
        End If
        iLastFrame = 10
    Case 11
        'Needs New Litter Frame 3
        Charcoal.iFrame = 10
        iLastFrame = 11
    Case 12
        'Needs Medicine Frame 1
        Charcoal.iFrame = 13
        iLastFrame = 12
    Case 13
        'Needs Medicine Frame 2
        If iLastFrame = 12 Then
            Charcoal.iFrame = 14
        Else
            Charcoal.iFrame = 12
        End If
        iLastFrame = 13
    Case 14
        'Needs Medicine Frame 3
        Charcoal.iFrame = 13
        iLastFrame = 14
    Case 15
        'Dead Frame 1
        Charcoal.iFrame = 16
    Case 16
        'Dead Frame 2
        Charcoal.iFrame = 17
    Case 17
        'Dead Frame 3
        MsgBox "Charcoal is dead!!!", vbOKOnly, "R.I.P."
  End Select
  i = BitBlt(picCharcoal.hDC, _
        0, 0, picCharcoal.ScaleWidth, picCharcoal.ScaleHeight, _
        frmGraphics.picCharcoal(Charcoal.iFrame).hDC, 0, 0, _
        SRCCOPY)
        
End Sub

Private Sub tmrNeeds_Timer()
  Dim iRand1, iRand2, i As Integer
  With Charcoal
    If .bNeedsFood = True Then
        .iHealth = .iHealth - 2
        .iHappiness = .iHappiness - 2
    End If
    If .bNeedsToPlay = True Then
        .iHappiness = .iHappiness - 2
    End If
    If .bNeedsLitter = True Then
        .iHealth = .iHealth - 4
        .iHappiness = .iHappiness - 2
    End If
    If .bIsSick = True Then
        .iHealth = .iHealth - 5
        .iHappiness = .iHappiness - 5
    End If
  End With
  iRand1 = Int((4 * Rnd) + 1)
  If iRand1 = 4 Then
      iRand2 = Int((4 * Rnd) + 1)
      Select Case iRand2
        Case 1
            'Charcoal needs food
            Charcoal.bNeedsFood = True
            Charcoal.iFrame = 3
        Case 2
            'Charcoal needs to play
            Charcoal.bNeedsToPlay = True
            Charcoal.iFrame = 6
        Case 3
            'Charcoal needs new litter
            Charcoal.bNeedsLitter = True
            Charcoal.iFrame = 9
        Case 4
            'Charcoal is sick
            Charcoal.bIsSick = True
            Charcoal.iFrame = 12
       End Select
       For i = 1 To 300
            Beep
       Next i
  End If
  If Charcoal.iHappiness <= 0 Or _
     Charcoal.iHealth <= 0 Then
        Charcoal.iFrame = 15
  End If
End Sub

Public Sub DetermineFrame()
  With Charcoal
    .iFrame = 0
    If .bNeedsFood = True Then .iFrame = 3
    If .bNeedsToPlay = True Then .iFrame = 6
    If .bNeedsLitter = True Then .iFrame = 9
    If .bIsSick = True Then .iFrame = 12
  End With
End Sub
