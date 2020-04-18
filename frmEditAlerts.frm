VERSION 5.00
Begin VB.Form frmEditAlerts 
   Caption         =   "Edit Alets"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   Icon            =   "frmEditAlerts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmEditAlerts.frx":6852
      Left            =   480
      List            =   "frmEditAlerts.frx":6862
      TabIndex        =   6
      Text            =   "Clicks"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox cboCondition 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmEditAlerts.frx":688A
      Left            =   2040
      List            =   "frmEditAlerts.frx":688C
      TabIndex        =   5
      Text            =   ">"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtAmount 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Text            =   "0"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.ListBox lstAlertList 
      Height          =   1815
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveChanges 
      Caption         =   "&Save Changes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblId 
      Caption         =   "AlertID:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblType 
      Caption         =   "Type:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblCondition 
      Caption         =   "Condition:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount:"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblAlertList 
      Alignment       =   2  'Center
      Caption         =   "Alert List"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmEditAlerts"
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
Dim num As Long
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveChanges_Click()
            If cboCondition.Text = ">" Then
                AlertList(num).ConditionType = 1
            End If
            If cboCondition.Text = "=" Then
                AlertList(num).ConditionType = 2
            End If
            
            If cboType.Text = "Clicks" Then
                AlertList(num).AlertType = 1
            End If
            If cboType.Text = "Impressions" Then
                AlertList(num).AlertType = 2
            End If
            If cboType.Text = "Earnings" Then
                AlertList(num).AlertType = 3
            End If
            If cboType.Text = "CTR" Then
                AlertList(num).AlertType = 4
            End If
        AlertList(num).Amount = txtAmount.Text
        MsgBox "Changes Saved for Alert Id: " & num, vbInformation
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Dim i As Long
    Dim strData As String
    
    cboCondition.AddItem ">"
    cboCondition.AddItem "="
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
Exit Sub
errHandle:

End Sub

Private Sub lstAlertList_Click()
On Error GoTo errHandle
    If lstAlertList.ListIndex <> -1 Then
        Dim Temp() As String
        
        Temp = Split(lstAlertList.Text, ":")
        num = CLng(Temp(0))
        txtAmount.Text = AlertList(num).Amount
        lblId.Caption = "AlertID: " & num
        If AlertList(num).ConditionType = 1 Then
            cboCondition.Text = ">"
        End If
        If AlertList(num).ConditionType = 2 Then
            cboCondition.Text = "="
        End If
        If AlertList(num).AlertType = 1 Then
            cboType.Text = "Clicks"
        End If
        If AlertList(num).AlertType = 2 Then
            cboType.Text = "Impressions"
        End If
        If AlertList(num).AlertType = 3 Then
            cboType.Text = "Earnings"
        End If
        If AlertList(num).AlertType = 4 Then
            cboType.Text = "CTR"
        End If
        txtAmount.Enabled = True
        Me.cboCondition.Enabled = True
        Me.cboType.Enabled = True
        cmdSaveChanges.Enabled = True
        'AlertList(num).ConditionType
    End If
Exit Sub
errHandle:

End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = 0
    If txtAmount.Text < 0 Then txtAmount.Text = 0
End Sub
