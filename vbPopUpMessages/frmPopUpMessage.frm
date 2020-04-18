VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPopUpMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   840
      ScaleHeight     =   2115
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   480
      Width           =   3375
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   225
         Left            =   600
         TabIndex        =   1
         Top             =   1470
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape shpBorder 
         Height          =   1185
         Left            =   540
         Top             =   570
         Width           =   2115
      End
      Begin VB.Image imgDown 
         Height          =   240
         Left            =   990
         Picture         =   "frmPopUpMessage.frx":0000
         Top             =   1170
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgUp 
         Height          =   240
         Left            =   780
         Picture         =   "frmPopUpMessage.frx":058A
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgClose 
         Height          =   225
         Left            =   2310
         Top             =   210
         Width           =   255
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblCaption"
         Height          =   195
         Left            =   870
         TabIndex        =   3
         Top             =   240
         Width           =   690
      End
      Begin VB.Image imgLogo 
         Height          =   240
         Left            =   570
         Top             =   210
         Width           =   240
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblMessage"
         Height          =   195
         Left            =   1230
         MouseIcon       =   "frmPopUpMessage.frx":0B14
         TabIndex        =   2
         Top             =   1110
         Width           =   825
      End
      Begin VB.Image imgBackground 
         Height          =   165
         Left            =   900
         Stretch         =   -1  'True
         Top             =   750
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmPopUpMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CloseGraceFully()
Public Event Closing()
Public Event Click()

Private menmDirection   As vbPopUpDirection

Public Sub AlignHolder()
    Select Case menmDirection
        Case vbPopUpDirection.vbPopUp
            picHolder.Align = 1
        Case vbPopUpDirection.vbPopDown
            picHolder.Align = 2
        Case vbPopUpDirection.vbPopLeft
            picHolder.Align = 3
        Case vbPopUpDirection.vbPopRight
            picHolder.Align = 4
    End Select
End Sub

Private Sub picHolder_Click()
    RaiseEvent CloseGraceFully
End Sub

Private Sub Form_Load()
    imgClose.Picture = imgUp.Picture
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Closing
End Sub

Private Sub imgBackground_Click()
    RaiseEvent CloseGraceFully
End Sub

Private Sub imgClose_Click()
    Unload Me
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgDown.Picture
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose.Picture = imgUp.Picture
End Sub

Private Sub lblMessage_Click()
    RaiseEvent Click
End Sub

Friend Property Let PopUpDirection(ByVal Value As vbPopUpDirection)
    menmDirection = Value
End Property

Public Sub ResizeControls(ByVal plngMsgHeight As Long, ByVal plngMsgWidth As Long)
Const BORDER_GAP As Long = 30
Const ICON_GAP As Long = 60
On Error GoTo ErrHandler
    With picHolder
        .Left = 0
        .Top = 0
        .Width = plngMsgWidth
        .Height = plngMsgHeight
    End With
    With imgBackground
        .Left = 0
        .Top = 0
        .Width = picHolder.ScaleWidth
        .Height = plngMsgHeight
    End With
    With imgLogo
        .Left = ICON_GAP
        .Top = ICON_GAP
    End With
    With imgClose
        .Left = picHolder.ScaleWidth - ICON_GAP - imgClose.Width
        .Top = ICON_GAP
    End With
    With shpBorder
        .Left = BORDER_GAP
        If imgLogo.Picture.Handle = 0 Then
            .Top = (2 * ICON_GAP) + lblCaption.Height
        Else
            .Top = (2 * ICON_GAP) + imgLogo.Height
        End If
        .Width = plngMsgWidth - (2 * BORDER_GAP)
        .Height = plngMsgHeight - .Top - BORDER_GAP
    End With
    With prgBar
        .Left = shpBorder.Left + ICON_GAP
        .Top = shpBorder.Top + shpBorder.Height - .Height - ICON_GAP
        .Width = shpBorder.Width - (2 * ICON_GAP)
    End With
    With lblCaption
        If imgLogo.Picture.Handle = 0 Then
            .Left = ICON_GAP
        Else
            .Left = (2 * ICON_GAP) + imgLogo.Width
        End If
        .Top = (shpBorder.Top - .Height) / 2
    End With
    With lblMessage
        .Left = (shpBorder.Width - .Width) / 2 + shpBorder.Left
        .Top = (shpBorder.Height - IIf(prgBar.Tag = True, prgBar.Height + (2 * ICON_GAP), 0) - .Height) / 2 + shpBorder.Top
    End With
    Exit Sub
ErrHandler:

End Sub
