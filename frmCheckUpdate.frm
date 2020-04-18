VERSION 5.00
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Checking for Update..."
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmCheckUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmCheckUpdate.frx":6852
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label lblVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmCheckUpdate"
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
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo errHandle

    lblVersion.Caption = "Current Version: " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Show
    Me.Refresh
    Me.Show
    If InternetCheckConnection("http://www.adsensealert.com/version.txt", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        Me.Caption = "Connection Failed! Check connection to internet!"
        lblMessage.ForeColor = vbRed
        lblMessage.FontBold = True
        lblMessage.Caption = Me.Caption
        
    Else
        
        Dim strData As String
        'strdata = GetUrl("http://www.adsensealert.com/version.txt", False)

          strData = GetUrl("http://www.adsensealert.com/version.txt")

       
        Dim Temp() As String
        Temp = Split(strData, ":")
        If UBound(Temp) <> 1 Then Exit Sub
        
        Dim strCurrent As String
        strCurrent = App.Major & "." & App.Minor & "." & App.Revision
        If strCurrent = Temp(1) Then
            lblMessage.FontBold = True
            lblMessage.ForeColor = vbGreen
            lblMessage.Caption = "You have the latest version!"
        Else
            lblMessage.FontBold = True
            lblMessage.ForeColor = vbRed
            lblMessage.Caption = "There is an update for this product!"
        End If
    End If
    
    Me.cmdClose.Enabled = True
Exit Sub
errHandle:
    MsgBox "Error_frmCheckUpdate_Load: " & Err.Description
End Sub
