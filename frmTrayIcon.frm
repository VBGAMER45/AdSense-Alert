VERSION 5.00
Begin VB.Form frmTrayIcon 
   Caption         =   "Tray Icon"
   ClientHeight    =   30
   ClientLeft      =   -945
   ClientTop       =   735
   ClientWidth     =   2340
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   30
   ScaleWidth      =   2340
   Visible         =   0   'False
   Begin VB.Menu mnuMenu 
      Caption         =   "File"
      Begin VB.Menu mnuTrial 
         Caption         =   "Buy Me!"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHideAdsenseAlert 
         Caption         =   "Hide Adsense Alert"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock Adsense Alert"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
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
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONDBLCLK = &H203


Private Sub Form_Load()

    
    #If Trial = True Then
        Me.mnuTrial.Visible = True
    #End If
    Me.Hide
    MySysTray.PopUpMessage = "Clicks: " & TodayClicks & " Impressions: " & TodayImpressions & " Earnings: " & TodayEarnings
    MySysTray.Initialize Me.hWnd, Me.icon, MySysTray.PopUpMessage
    MySysTray.ShowIcon
    Me.Hide
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim msgCallBackMessage As Long
  msgCallBackMessage = X / Screen.TwipsPerPixelX
  Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
      MySysTray.TipText = MySysTray.PopUpMessage
   Case WM_RBUTTONDOWN
        Me.PopupMenu mnuMenu
   Case WM_LBUTTONDBLCLK
     If IsLocked = True Then
        frmEnterPassword.Tag = "main"
        frmEnterPassword.Show
        Exit Sub
     End If
        frmMain.Show
        IsInTray = False

   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MySysTray.HideIcon
    End
End Sub



Private Sub mnuExit_Click()
    Dim Response As VbMsgBoxResult
    Response = MsgBox("Are you sure you want to quit?", vbYesNo + vbInformation, "Quit Adsense Alert?")
    If Response = vbYes Then
        Call frmMain.SaveConfigInformation
        Unload frmOptions
        Unload frmSetPassword
        Unload frmAbout
        Unload frmMain
        Unload Me
        End
    End If
End Sub

Private Sub mnuFileLoginAdsense_Click()
    On Error Resume Next

     ShellExecute Me.hWnd, vbNullString, "https://www.google.com/adsense/login.do?username=" & UsernameSafe & "&password=" & PasswordSafe, vbNullString, "C:\", SW_SHOWNORMAL
     
End Sub

Private Sub mnuFileUpdateTodayStats_Click()
    If bLoggedIn = True Then
        Call modAdsense.GetTodaysAdData(frmMain.lstAdReport)
    Else
        MsgBox "You are not logged in!", vbExclamation
    End If
End Sub

Private Sub mnuHideAdsenseAlert_Click()
    If IsLocked = True Then
        frmEnterPassword.Tag = "main"
        frmEnterPassword.Show
        Exit Sub
    End If
    If mnuHideAdsenseAlert.Caption = "Hide Adsense Alert" Then
        Unload frmSetPassword
        Unload frmAbout
        Unload frmEnterPassword
        Unload frmOptions
        frmMain.Hide
        IsInTray = True
        frmTrayIcon.mnuHideAdsenseAlert.Caption = "Show Adsense Alert"
    Else
        frmMain.Show
        IsInTray = False
        frmTrayIcon.mnuHideAdsenseAlert.Caption = "Hide Adsense Alert"

    End If
    
End Sub

Private Sub mnuLock_Click()
    If AdsenseAlertPassword = "" Then
        MsgBox "No Password has been set. Please set a password.", vbInformation
        frmSetPassword.Show
        Exit Sub
    Else
        frmMain.Hide
        frmOptions.Hide
        frmAbout.Hide
        frmSetPassword.Hide
        frmEnterPassword.Hide
        frmTrayIcon.mnuHideAdsenseAlert.Caption = "Show Adsense Alert"
        IsLocked = True
        
    End If
End Sub

Private Sub mnuOptions_Click()
    If IsLocked = True Then
        frmEnterPassword.Tag = "options"
        frmEnterPassword.Show
        Exit Sub
    End If
    frmOptions.Show
End Sub

Private Sub mnuTrial_Click()
    On Error Resume Next
     ShellExecute Me.hWnd, vbNullString, "http://www.adsensealert.com/order.php", vbNullString, "C:\", SW_SHOWNORMAL
     
End Sub

