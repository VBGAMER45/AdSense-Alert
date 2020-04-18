VERSION 5.00
Begin VB.Form frmEnterPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "frmEnterPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter the password for Adsense Alert"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image imgLock 
      Height          =   720
      Left            =   120
      Picture         =   "frmEnterPassword.frx":6852
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmEnterPassword"
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
    If txtPassword.Text = "" Then
        MsgBox "You need to enter a password!", vbExclamation
        Exit Sub
    End If
    
    If txtPassword.Text = AdsenseAlertPassword Then
        IsLocked = False
        MsgBox "Unlocked access granted", vbInformation
    Else
        MsgBox "Invalid Password!", vbCritical
    End If
    Unload Me
End Sub
