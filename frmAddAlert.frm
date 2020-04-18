VERSION 5.00
Begin VB.Form frmAddAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Alert"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAddAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Text            =   "0"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox cboCondition 
      Height          =   315
      ItemData        =   "frmAddAlert.frx":6852
      Left            =   1680
      List            =   "frmAddAlert.frx":6854
      TabIndex        =   7
      Text            =   ">"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      Picture         =   "frmAddAlert.frx":6856
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddAlert 
      Caption         =   "&Add Alert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      Picture         =   "frmAddAlert.frx":D0A8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmAddAlert.frx":138FA
      Left            =   120
      List            =   "frmAddAlert.frx":1390A
      TabIndex        =   1
      Text            =   "Clicks"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount:"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblCondition 
      Caption         =   "Condition:"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblType 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblAdPerformace 
      Caption         =   "Ad Performance"
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
      Left            =   1613
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   "Here you can setup alerts when certain conditions are met."
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblAlert 
      Caption         =   "Add Alert's"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image imgAlert 
      Height          =   720
      Left            =   120
      Picture         =   "frmAddAlert.frx":13932
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmAddAlert"
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

Private Sub cmdAddAlert_Click()
    #If Trial = True Then
        If UBound(AlertList) >= 1 Then
            MsgBox "Trial Version only allows one alert!", vbCritical
        End If
    #End If
        
    If txtAmount.Text = 0 Then
        MsgBox "You need to an amount other than zero!", vbExclamation
        Exit Sub
    End If
    
    #If Trial = True Then
        If UBound(AlertList) >= 1 Then
            Exit Sub
        End If
    #End If
    
    'Check if there is an empty spot
    Dim i As Long
    Dim Found As Boolean
    Found = False
    For i = 0 To UBound(AlertList)
        If AlertList(i).Amount = 0 Then
            If cboCondition.Text = ">" Then
                AlertList(i).ConditionType = 1
            End If
            If cboCondition.Text = "=" Then
                AlertList(i).ConditionType = 2
            End If
            
            If cboType.Text = "Clicks" Then
                AlertList(i).AlertType = 1
            End If
            If cboType.Text = "Impressions" Then
                AlertList(i).AlertType = 2
            End If
            If cboType.Text = "Earnings" Then
                AlertList(i).AlertType = 3
            End If
            If cboType.Text = "CTR" Then
                AlertList(i).AlertType = 4
            End If
            AlertList(i).AlertOn = True
            AlertList(i).Amount = txtAmount.Text
            AlertList(i).AlertDate = Date
            
            Found = True
            Exit For
        End If
    Next
    
    If Found = False Then
        If cboCondition.Text = ">" Then
            AlertList(UBound(AlertList)).ConditionType = 1
        End If
        If cboCondition.Text = "=" Then
            AlertList(UBound(AlertList)).ConditionType = 2
        End If
        
        If cboType.Text = "Clicks" Then
            AlertList(UBound(AlertList)).AlertType = 1
        End If
        If cboType.Text = "Impressions" Then
            AlertList(UBound(AlertList)).AlertType = 2
        End If
        If cboType.Text = "Earnings" Then
            AlertList(UBound(AlertList)).AlertType = 3
        End If
        If cboType.Text = "CTR" Then
            AlertList(UBound(AlertList)).AlertType = 4
        End If
        AlertList(UBound(AlertList)).AlertOn = True
        AlertList(UBound(AlertList)).Amount = txtAmount.Text
        AlertList(UBound(AlertList)).AlertDate = Date
        ReDim Preserve AlertList(UBound(AlertList) + 1)
    End If
    MsgBox "Alert Added!", vbInformation
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    cboCondition.AddItem ">"
    cboCondition.AddItem "="
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = 0
    If txtAmount.Text < 0 Then txtAmount.Text = 0
End Sub
