VERSION 5.00
Begin VB.Form frmSelectChannel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channel Select"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Done"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ListBox lstChannels 
      Height          =   2085
      Left            =   480
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Channel List"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblUnSelectAll 
      Caption         =   "Unselect All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblSelectAll 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo errHandle
    Dim i As Long, j As Long
    For i = 0 To lstChannels.ListCount - 1
        If lstChannels.Selected(i) = True Then
            For j = 0 To UBound(ChannelList)
                If ChannelList(j).ChannelName = lstChannels.List(i) Then
                    ChannelList(j).Selected = True
                End If
            Next
        Else
            For j = 0 To UBound(ChannelList)
                If ChannelList(j).ChannelName = lstChannels.List(i) Then
                    ChannelList(j).Selected = False
                End If
            Next
        End If
    Next
    
    Unload Me
Exit Sub
errHandle:

End Sub

Private Sub cmdUpdate_Click()
On Error GoTo errHandle
    Dim strData As String
       Dim h As HTTPClass
    
       Set h = New HTTPClass


       If h.OpenHTTP("www.google.com", INTERNET_DEFAULT_HTTPS_PORT) Then
          strData = h.SendRequest("/adsense/login.do?destination=/adsense/report/aggregate%3Fproduct%3Dafc&username=" & UserName & "&password=" & Password, "GET")
       End If
        Call GetUrlChannels(strData)
    Set h = Nothing

    lstChannels.Clear
    Dim i As Long
    For i = 0 To UBound(ChannelList)
        If ChannelList(i).ChannelName <> "" Then
            lstChannels.AddItem ChannelList(i).ChannelName
            If ChannelList(i).Selected = True Then
                lstChannels.Selected(lstChannels.ListCount - 1) = True
            End If
        End If
    Next
Exit Sub
errHandle:
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    lstChannels.Clear
    Dim i As Long
    For i = 0 To UBound(ChannelList)
        If ChannelList(i).ChannelName <> "" Then
            lstChannels.AddItem ChannelList(i).ChannelName
            If ChannelList(i).Selected = True Then
                lstChannels.Selected(lstChannels.ListCount - 1) = True
            End If
        End If
    Next
Exit Sub
errHandle:
    
End Sub

Private Sub lblSelectAll_Click()
    On Error GoTo errHandle
        Dim i As Long
        For i = 0 To lstChannels.ListCount - 1
            lstChannels.Selected(i) = True
        Next
    Exit Sub
errHandle:
End Sub

Private Sub lblUnSelectAll_Click()
    On Error GoTo errHandle
        Dim i As Long
        For i = 0 To lstChannels.ListCount - 1
            lstChannels.Selected(i) = False
        Next
    Exit Sub
errHandle:
End Sub

