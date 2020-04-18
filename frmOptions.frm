VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEmailAlerts 
      Caption         =   "Email Alerts"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CheckBox chkMsn 
      Caption         =   "Use Msn Style updates"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Sound On Update"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtEmailPath 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "http://www.site.com/mail.php"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CheckBox chkEmailStats 
      Caption         =   "Email Adsense Updates"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "10"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CheckBox chkOnStartUp 
      Caption         =   "Run on Start Up"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "If you want to use mailing of adsense updates copy the mail.php to your website"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Path To PHP script:"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblEmailOptions 
      Caption         =   "Email Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Image imgEmailOptions 
      Height          =   720
      Left            =   120
      Picture         =   "frmOptions.frx":6852
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label lblIn 
      Caption         =   "Inveral in minutes:"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblUpdate 
      Caption         =   "Update Interval"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmOptions.frx":D0A4
      Top             =   2040
      Width           =   720
   End
   Begin VB.Image imgOptions 
      Height          =   720
      Left            =   120
      Picture         =   "frmOptions.frx":138F6
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'AdSense Alert
'VisualBasicZone.com 2005
'Jonathan Valentin
'********************************************
Option Explicit
#Const Trial = False

Private Sub chkEmailAlerts_Click()
    If chkEmailAlerts.Value = vbChecked Then
        EmailAlerts = True
    Else
        EmailAlerts = False
    End If
End Sub

Private Sub chkEmailStats_Click()
    If chkEmailStats.Value = vbChecked Then
        EmailUpdates = True
    Else
        EmailUpdates = False
    End If
End Sub


Private Sub chkOnStartUp_Click()
    If chkOnStartUp.Value = vbChecked Then
        Call modGlobals.RegRun(App.Path & "\AdsenseAlert.exe tray", "AdsenseAlert")
        RunOnStartUp = True
    Else
        Call modGlobals.RemoveRegRun("AdsenseAlert")
        RunOnStartUp = False
        
    End If
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtInterval.Text = UpdateInterval
    Me.txtEmailPath.Text = EmailUrl
    If RunOnStartUp = True Then
        chkOnStartUp.Value = vbChecked
    Else
        chkOnStartUp.Value = vbUnchecked
    End If
    If EmailAlerts = True Then
        Me.chkEmailAlerts.Value = vbChecked
    Else
        chkEmailAlerts.Value = vbUnchecked
    End If
    If EmailUpdates = True Then
        Me.chkEmailStats.Value = vbChecked
    Else
        chkEmailStats.Value = vbUnchecked
    End If
    If UseMsnUpdates = True Then
        chkMsn.Value = vbChecked
    Else
        chkMsn.Value = vbUnchecked
    End If
    If SoundOnUpdate = True Then
        chkSound.Value = vbChecked
    Else
        chkSound.Value = vbUnchecked
    End If
End Sub

Private Sub txtEmailPath_Change()
    EmailUrl = txtEmailPath.Text
End Sub

Private Sub txtInterval_Change()
On Error GoTo errHandle
    If IsNumeric(txtInterval.Text) = False Then txtInterval.Text = "10"
    If txtInterval.Text <= 0 Then txtInterval.Text = "10"
    UpdateInterval = txtInterval.Text
    If UpdateInterval > CurrentUpdate Then CurrentUpdate = UpdateInterval
Exit Sub
errHandle:
    MsgBox "frmOptions_txtInterval_Change: " & Err.Description
End Sub
