VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PopUp Messages Demo"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFading 
      Caption         =   "Fading"
      Height          =   255
      Left            =   630
      TabIndex        =   27
      Top             =   3090
      Width           =   1215
   End
   Begin VB.Frame fraLogos 
      Caption         =   "Logos"
      Height          =   1005
      Left            =   3720
      TabIndex        =   3
      Top             =   330
      Width           =   1905
      Begin VB.OptionButton optLogo 
         Caption         =   "One"
         Height          =   405
         Index           =   0
         Left            =   660
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optLogo 
         Caption         =   "Two"
         Height          =   405
         Index           =   1
         Left            =   660
         TabIndex        =   4
         Top             =   540
         Width           =   1005
      End
      Begin VB.Image imgLogo 
         Height          =   240
         Index           =   0
         Left            =   210
         Picture         =   "frmMain.frx":0000
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgLogo 
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":058A
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Frame fraWavs 
      Caption         =   "Wavs"
      Height          =   1545
      Left            =   3720
      TabIndex        =   0
      Top             =   1440
      Width           =   1905
      Begin VB.OptionButton optWav 
         Caption         =   "Online"
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   17
         Top             =   1020
         Width           =   1005
      End
      Begin VB.OptionButton optWav 
         Caption         =   "Typing"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optWav 
         Caption         =   "Email"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   1605
      Left            =   2520
      TabIndex        =   22
      Top             =   2700
      Width           =   1905
      Begin VB.OptionButton optDirection 
         Caption         =   "Right"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   26
         Top             =   1260
         Width           =   1095
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Left"
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   25
         Top             =   990
         Width           =   1095
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Down"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Up"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   23
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkSticky 
      Caption         =   "Sticky"
      Height          =   255
      Left            =   630
      TabIndex        =   21
      Top             =   3390
      Width           =   1215
   End
   Begin VB.CheckBox chkUseParent 
      Caption         =   "Use PictureBox"
      Height          =   255
      Left            =   630
      TabIndex        =   20
      Top             =   3990
      Width           =   1425
   End
   Begin VB.CheckBox chkAdd 
      Caption         =   "Auto Add"
      Height          =   255
      Left            =   630
      TabIndex        =   19
      Top             =   3690
      Width           =   1215
   End
   Begin VB.Timer tmrAdd 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2220
      Top             =   5280
   End
   Begin VB.PictureBox picHolder 
      Height          =   4995
      Left            =   5790
      ScaleHeight     =   4935
      ScaleWidth      =   5535
      TabIndex        =   18
      Top             =   180
      Width           =   5595
   End
   Begin VB.Timer tmrWorking 
      Interval        =   100
      Left            =   780
      Top             =   5280
   End
   Begin VB.HScrollBar scrScrollDelay 
      Height          =   225
      LargeChange     =   10
      Left            =   1800
      Max             =   200
      Min             =   1
      TabIndex        =   13
      Top             =   4920
      Value           =   1
      Width           =   3885
   End
   Begin VB.HScrollBar scrMovementIndex 
      Height          =   225
      LargeChange     =   5
      Left            =   1800
      Max             =   100
      Min             =   1
      TabIndex        =   12
      Top             =   4650
      Value           =   1
      Width           =   3885
   End
   Begin VB.HScrollBar scrShowDelay 
      Height          =   225
      LargeChange     =   100
      Left            =   1800
      Max             =   4000
      Min             =   1
      SmallChange     =   5
      TabIndex        =   11
      Top             =   4380
      Value           =   1
      Width           =   3885
   End
   Begin VB.CommandButton cmdShowDefault 
      Caption         =   "Show Default"
      Height          =   405
      Left            =   2340
      TabIndex        =   10
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show PopUp"
      Height          =   405
      Left            =   1110
      TabIndex        =   9
      Top             =   180
      Width           =   1125
   End
   Begin VB.Frame fraBackGrounds 
      Caption         =   "BackGrounds"
      Height          =   2175
      Left            =   270
      TabIndex        =   6
      Top             =   720
      Width           =   3765
      Begin VB.OptionButton optBackGround 
         Caption         =   "One"
         Height          =   225
         Index           =   0
         Left            =   780
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optBackGround 
         Caption         =   "Two"
         Height          =   225
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.Image imgBack 
         BorderStyle     =   1  'Fixed Single
         Height          =   1050
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":0B14
         Stretch         =   -1  'True
         Top             =   630
         Width           =   1440
      End
      Begin VB.Image imgBack 
         BorderStyle     =   1  'Fixed Single
         Height          =   1080
         Index           =   1
         Left            =   1800
         Picture         =   "frmMain.frx":22E7
         Stretch         =   -1  'True
         Top             =   630
         Width           =   1410
      End
   End
   Begin VB.Image imgSticky 
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Left            =   270
      Picture         =   "frmMain.frx":2848
      Stretch         =   -1  'True
      Top             =   5940
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "Scroll Delay:"
      Height          =   195
      Left            =   300
      TabIndex        =   16
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Movement Index:"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   4650
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Show Delay:"
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   4380
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjPopups   As PopUpMessages
Attribute mobjPopups.VB_VarHelpID = -1

Private mobjDefault             As PopUpMessage

Private Sub chkAdd_Click()
    tmrAdd.Enabled = (chkAdd.Value = vbChecked)
End Sub

Private Sub chkFading_Click()
    mobjPopups.AllowFading = (chkFading.Value = vbChecked)
End Sub

Private Sub chkUseParent_Click()
    SetParent chkUseParent.Value = vbChecked
End Sub

Private Sub cmdShowDefault_Click()
    mobjPopups.Show mobjDefault
    tmrWorking.Enabled = True
End Sub

Private Sub Form_Load()
    Set mobjPopups = New PopUpMessages
    With mobjPopups
       ' .XPos = Screen.Width / 2
       ' .YPos = 0
       ' .PopUpDirection = vbPopDown
        scrShowDelay.Value = .ShowDelay
        scrMovementIndex.Value = .MovementIndex
        scrScrollDelay.Value = .ScrollDelay
    End With
    SetupDefaultPopup
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjPopups = Nothing
    Set mobjDefault = Nothing
End Sub

Private Sub cmdShow_Click()
    AddPopup
End Sub

Private Sub mobjPopups_PopupMessageClicked(Item As PopUpMessage)
    'MsgBox Item.Message & " Clicked"
End Sub

Private Sub SetupDefaultPopup()
    Set mobjDefault = New PopUpMessage
    With mobjDefault
        Set .Background = imgBack.Item(1)
        .ForeColor = vbWhite
        Set .Logo = imgLogo.Item(1)
        .WavFile = App.Path & "\resources\sounds\newemail.wav"
        .Caption = "New Email"
        .Message = "You have received" & vbCrLf & "4 new emails." & vbCrLf & "Downloading..."
        .Clickable = True
        .ProgressBar = True
    End With
End Sub

Private Sub picSticky_Click()

End Sub

Private Sub optDirection_Click(Index As Integer)
    mobjPopups.PopUpDirection = Index
    SetParent (mobjPopups.ParentHandle <> 0)
End Sub

Private Sub scrShowDelay_Change()
    mobjPopups.ShowDelay = scrShowDelay.Value
End Sub

Private Sub scrMovementIndex_Change()
    mobjPopups.MovementIndex = scrMovementIndex.Value
End Sub

Private Sub scrScrollDelay_Change()
    mobjPopups.ScrollDelay = scrScrollDelay.Value
End Sub

Private Sub tmrAdd_Timer()
    AddPopup
End Sub

Private Sub tmrWorking_Timer()
    If mobjDefault.Value = 100 Then
        mobjDefault.Value = 1
    End If
    mobjDefault.Value = mobjDefault.Value + 1
    If mobjDefault.Value = 100 Or Not mobjDefault.Visible Then
        tmrWorking.Enabled = False
    End If
End Sub

Private Sub AddPopup()
Dim objPopUp    As PopUpMessage
    Set objPopUp = New PopUpMessage
    With objPopUp
        .Caption = "Wokawidget Software"
        .Message = "I am a doctor, it is " & vbCrLf & "necrotising fasciitis says:" & vbCrLf & "'Woof...'"
        .Clickable = False
        .Sticky = (chkSticky.Value = vbChecked)
        If chkSticky.Value = vbChecked Then
            Set .Background = imgSticky
        Else
            If optBackGround.Item(0).Value Then
                Set .Background = imgBack.Item(0)
            Else
                Set .Background = imgBack.Item(1)
            End If
        End If
        If optLogo.Item(0).Value Then
            Set .Logo = imgLogo.Item(0)
        Else
            Set .Logo = imgLogo.Item(1)
        End If
        If optWav.Item(0).Value Then
            .WavFile = App.Path & "\resources\sounds\type.wav"
        ElseIf optWav.Item(1).Value Then
            .WavFile = App.Path & "\resources\sounds\newemail.wav"
        ElseIf optWav.Item(2).Value Then
            .WavFile = App.Path & "\resources\sounds\online.wav"
        End If
        
    End With
    mobjPopups.Show objPopUp
End Sub

Private Sub SetParent(ByVal pblnUsePicBox As Long)
    With mobjPopups
        .ParentHandle = IIf(pblnUsePicBox, picHolder.hWnd, 0)
        Select Case mobjPopups.PopUpDirection
            Case vbPopUpDirection.vbPopUp
                If pblnUsePicBox Then
                    .XPos = 0
                    .YPos = picHolder.ScaleHeight
                Else
                    .XPos = GetDesktopWidth - .MessageWidth
                    .YPos = GetDesktopHeight
                End If
            Case vbPopUpDirection.vbPopDown
                If pblnUsePicBox Then
                    .XPos = 0
                    .YPos = 0
                Else
                    .XPos = GetDesktopWidth - .MessageWidth
                    .YPos = 0
                End If
            Case vbPopUpDirection.vbPopLeft
                If pblnUsePicBox Then
                    .XPos = picHolder.ScaleWidth
                    .YPos = 0
                Else
                    .XPos = GetDesktopWidth
                    .YPos = 0
                End If
            Case vbPopUpDirection.vbPopRight
                If pblnUsePicBox Then
                    .XPos = 0
                    .YPos = 0
                Else
                    .XPos = 0
                    .YPos = 0
                End If
        End Select
    End With
End Sub
