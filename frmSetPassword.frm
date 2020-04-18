VERSION 5.00
Begin VB.Form frmSetPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Password"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSetPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtPassword2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtPassword1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblOldPassword 
      Caption         =   "Old Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblConfrim 
      Caption         =   "Confrim:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblSetPassword 
      Caption         =   "Set Password used to access Adsense Alert"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image imgLock 
      Height          =   720
      Left            =   120
      Picture         =   "frmSetPassword.frx":6852
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmSetPassword"
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
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtPassword1.Text = "" Or txtPassword2.Text = "" Then
        MsgBox "You need to enter a password!", vbInformation
        Exit Sub
    End If
    If txtPassword1.Text <> txtPassword2.Text Then
        MsgBox "Passwords do not match!", vbCritical
        Exit Sub
    End If
    If AdsenseAlertPassword <> "" Then
        If txtOldPassword.Text <> AdsenseAlertPassword Then
            MsgBox "Your old password doesn't match the one on file!", vbCritical, "Wrong Password!"
            Exit Sub
        End If
    End If
    
    AdsenseAlertPassword = txtPassword1.Text
    
    MsgBox "Password was set.  To Lock program click the x button on the main form and in the tray goto lock adsense alert", vbInformation
    Unload Me
End Sub

Private Sub Form_Load()
    If AdsenseAlertPassword <> "" Then
        lblOldPassword.Visible = True
        txtOldPassword.Visible = True
    End If
End Sub


