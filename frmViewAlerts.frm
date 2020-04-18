VERSION 5.00
Begin VB.Form frmViewAlerts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Alerts"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmViewAlerts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Alert"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ListBox lstAlertList 
      Height          =   1620
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmViewAlerts.frx":6852
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblView 
      Caption         =   "View Alert List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmViewAlerts"
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
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
On Error GoTo errHandle
    Dim Temp() As String
    Dim num As Long
    Temp = Split(lstAlertList.Text, ":")
    num = CLng(Temp(0))
    AlertList(num).Amount = 0
    AlertList(num).AlertOn = False
    
    Me.lstAlertList.RemoveItem lstAlertList.ListIndex
    cmdRemove.Enabled = False
Exit Sub
errHandle:
    MsgBox "Error_frmViewAlerts_cmdRemove: " & Err.Description
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strData As String
    For i = 0 To UBound(AlertList)
        
        If AlertList(i).Amount <> 0 Then
            strData = i & ": "
            If AlertList(i).AlertType = 1 Then
                strData = strData & "Clicks "
            ElseIf AlertList(i).AlertType = 2 Then
                strData = strData & "Impressions "
            ElseIf AlertList(i).AlertType = 3 Then
                strData = strData & "Earnings "
            ElseIf AlertList(i).AlertType = 4 Then
                strData = strData & "CTR "
            End If
            If AlertList(i).ConditionType = 1 Then
                strData = strData & ">"
            End If
            If AlertList(i).ConditionType = 2 Then
                strData = strData & "="
            End If
            strData = strData & " " & AlertList(i).Amount
            lstAlertList.AddItem strData
        End If
    Next
End Sub

Private Sub lstAlertList_Click()
    If lstAlertList.ListIndex <> -1 Then
        cmdRemove.Enabled = True
    End If
End Sub
