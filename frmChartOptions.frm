VERSION 5.00
Begin VB.Form frmChartOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart Options"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmChartOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboChartType 
      Height          =   315
      ItemData        =   "frmChartOptions.frx":6852
      Left            =   360
      List            =   "frmChartOptions.frx":687A
      TabIndex        =   10
      Text            =   "2dBar"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox txtyAxis 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtXAxis 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CheckBox chkShowLegend 
      Caption         =   "Show Legend"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveChanges 
      Caption         =   "&Save Changes"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblChartType 
      Caption         =   "Chart Type"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label lblYaxis 
      Caption         =   "Y Axis"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblXaxis 
      Caption         =   "X Axis"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblGraph 
      Caption         =   "Graph Title"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmChartOptions"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveChanges_Click()
    With frmGraph.Chart
        .TitleText = txtTitle.Text
        If chkShowLegend.Value = vbChecked Then
            .ShowLegend = True
        Else
            .ShowLegend = False
        End If
        
        Select Case cboChartType.Text
            
            Case "2dArea"
                .chartType = VtChChartType2dArea
            Case "2dBar"
                .chartType = VtChChartType2dBar
            Case "2dCombination"
                .chartType = VtChChartType2dCombination
            Case "2dLine"
                .chartType = VtChChartType2dLine
            Case "2dPie"
                .chartType = VtChChartType2dPie
            Case "2dStep"
                .chartType = VtChChartType2dStep
            Case "2dXY"
                .chartType = VtChChartType2dXY
            Case "3dArea"
                .chartType = VtChChartType3dArea
            Case "3dBar"
                .chartType = VtChChartType3dBar
            Case "3dCombination"
                .chartType = VtChChartType3dCombination
            Case "3dLine"
                .chartType = VtChChartType3dLine
            Case "3dStep"
                .chartType = VtChChartType3dStep
                
        End Select
        frmGraph.Tag = cboChartType.Text
    
    End With
    
    
    With frmGraph.Chart.Plot
        .Axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = txtXAxis.Text
        .Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = txtyAxis.Text
    End With
End Sub
