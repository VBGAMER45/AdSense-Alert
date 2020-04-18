VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraph 
   Caption         =   "Graph"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmGraph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboReport 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   5040
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   975
      Left            =   480
      Picture         =   "frmGraph.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   975
      Left            =   5280
      Picture         =   "frmGraph.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Chart Options"
      Height          =   975
      Left            =   2880
      Picture         =   "frmGraph.frx":138F6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4680
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSChart20Lib.MSChart Chart 
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "frmGraph.frx":1A148
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmGraph.frx":1C49E
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label lblReportOptions 
      Caption         =   "Reports:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
End
Attribute VB_Name = "frmGraph"
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

Private Sub cboReport_Change()
    If frmGraph.Chart.Tag = "adreport" Then
        Select Case cboReport.Text
            Case "Impressions"
                  Call frmGraph.SetUpChartAdReport("Adsense Alert Impressions Report", "Date", "Impressions", 1)
            Case "Clicks"
                Call frmGraph.SetUpChartAdReport("Adsense Alert Clicks Report", "Date", "Clicks", 2)
            Case "CTR"
                Call frmGraph.SetUpChartAdReport("Adsense Alert CTR Report", "Date", "CTR", 3)
            Case "CPM"
                Call frmGraph.SetUpChartAdReport("Adsense Alert CPM Report", "Date", "CPM", 4)
            Case "Earnings"
                Call frmGraph.SetUpChartAdReport("Adsense Alert Earnings Report", "Date", "Earnings", 5)
        End Select
    ElseIf frmGraph.Chart.Tag = "searchreport" Then
        Select Case cboReport.Text
            Case "Impressions"
                  Call frmGraph.SetUpChartSearchReport("Adsense Alert Impressions Report", "Date", "Impressions", 1)
            Case "Clicks"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert Clicks Report", "Date", "Clicks", 2)
            Case "CTR"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert CTR Report", "Date", "CTR", 3)
            Case "CPM"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert CPM Report", "Date", "CPM", 4)
            Case "Earnings"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert Earnings Report", "Date", "Earnings", 5)
        End Select
    ElseIf frmGraph.Chart.Tag = "payreport" Then
        Call frmGraph.SetUpChartSearchReport("Adsense Alert Payment Report", "Date", "Earnings", 2)
    End If
End Sub

Private Sub cboReport_Click()
    If frmGraph.Chart.Tag = "adreport" Then
        Select Case cboReport.Text
            Case "Impressions"
                  Call frmGraph.SetUpChartAdReport("Adsense Alert Impressions Report", "Date", "Impressions", 1)
            Case "Clicks"
                Call frmGraph.SetUpChartAdReport("Adsense Alert Clicks Report", "Date", "Clicks", 2)
            Case "CTR"
                Call frmGraph.SetUpChartAdReport("Adsense Alert CTR Report", "Date", "CTR", 3)
            Case "CPM"
                Call frmGraph.SetUpChartAdReport("Adsense Alert CPM Report", "Date", "CPM", 4)
            Case "Earnings"
                Call frmGraph.SetUpChartAdReport("Adsense Alert Earnings Report", "Date", "Earnings", 5)
        End Select
    ElseIf frmGraph.Chart.Tag = "searchreport" Then
        Select Case cboReport.Text
            Case "Impressions"
                  Call frmGraph.SetUpChartSearchReport("Adsense Alert Impressions Report", "Date", "Impressions", 1)
            Case "Clicks"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert Clicks Report", "Date", "Clicks", 2)
            Case "CTR"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert CTR Report", "Date", "CTR", 3)
            Case "CPM"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert CPM Report", "Date", "CPM", 4)
            Case "Earnings"
                Call frmGraph.SetUpChartSearchReport("Adsense Alert Earnings Report", "Date", "Earnings", 5)
        End Select
    ElseIf frmGraph.Chart.Tag = "payreport" Then
        Call frmGraph.SetUpChartSearchReport("Adsense Alert Payment Report", "Date", "Earnings", 2)
    End If
End Sub

Private Sub Chart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Part As Integer, Series As Integer, DataPoint As Integer
    Dim Index3 As Integer, Index4 As Integer
    Dim oValue As Double, NullFlag As Integer
Exit Sub
    With Chart

        
        .TwipsToChartPart X, Y, Part, Series, DataPoint, Index3, Index4
    

        
        'Show value in ToolTipText when select point
        If Part = VtChPartTypeChart Then
            .ToolTipText = "Chart Area"
        ElseIf Part = VtChPartTypeTitle Then
            .ToolTipText = "Chart Title"
        ElseIf Part = VtChPartTypeFootnote Then
            .ToolTipText = "Footnote"
        ElseIf Part = VtChPartTypeLegend Then
            .ToolTipText = "Legend"
        ElseIf Part = VtChPartTypePlot Then
            .ToolTipText = "Plot Area"
        ElseIf Part = VtChPartTypePoint Or Part = VtChPartTypePointLabel Then 'Or VtChPartTypeSeries Then
            
            .DataGrid.GetData DataPoint, Series, oValue, NullFlag
            .ToolTipText = "Value := " & oValue
        ElseIf Part = VtChPartTypeSeries Then
             .DataGrid.GetData DataPoint + 1, Series, oValue, NullFlag
            .ToolTipText = "Value := " & oValue
        ElseIf Part = VtChPartTypeAxis Then
            .ToolTipText = "Plot Axis"
        ElseIf Part = VtChPartTypeAxisLabel Then
            .ToolTipText = "Axis Label"
        ElseIf Part = VtChPartTypeAxisTitle Then
            .ToolTipText = "Axis Title"
        Else
            
            .ToolTipText = ""
        End If
    End With
End Sub



Private Sub cmdOptions_Click()
    frmChartOptions.txtTitle = Chart.TitleText
    frmChartOptions.txtXAxis.Text = Chart.Plot.Axis(VtChAxisIdX).AxisTitle.Text
    frmChartOptions.txtyAxis.Text = Chart.Plot.Axis(VtChAxisIdY).AxisTitle.Text
    frmChartOptions.cboChartType.Text = Me.Tag
    If Chart.ShowLegend = True Then
        frmChartOptions.chkShowLegend.Value = vbChecked
    Else
        frmChartOptions.chkShowLegend.Value = vbUnchecked
    End If
    frmChartOptions.Show
End Sub


Private Sub cmdPrint_Click()
    #If Trial = True Then
        MsgBox "Not in Trial Version", vbExclamation
    #Else
        CD.DialogTitle = "Print Graph"
        CD.Flags = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection
        CD.ShowPrinter
        Clipboard.Clear
        Chart.EditCopy
        Dim i As Long
        For i = 1 To CD.Copies
            Call PrintChart
        Next
    #End If
End Sub
Public Sub SetUpChartAdReport(ChartTitle As String, xAxis As String, yAxis As String, ChartIndex As Integer)
    If frmMain.lstAdReport.ListItems.Count < 1 Then
        MsgBox "No data to create chart", vbInformation
        Exit Sub
    End If

    frmGraph.Chart.Tag = "adreport"
    frmGraph.Chart.Title.Text = ChartTitle
    frmGraph.Chart.Plot.Axis(VtChAxisIdX).AxisTitle.Text = xAxis
    frmGraph.Chart.Plot.Axis(VtChAxisIdY).AxisTitle.Text = yAxis
    
    Dim i As Long
    
    frmGraph.Chart.Plot.Wall.Brush.Style = VtBrushStyleSolid
    frmGraph.Chart.Plot.Wall.Brush.FillColor.Set 255, 255, 225
   

    frmGraph.Chart.RowCount = frmMain.lstAdReport.ListItems.Count - 2 'UBound(arrData) + 1
    frmGraph.Chart.ColumnCount = 1
 
    For i = 1 To frmGraph.Chart.RowCount
        
        frmGraph.Chart.Row = i
        frmGraph.Chart.RowLabel = frmMain.lstAdReport.ListItems.Item(i).Text

       Call frmGraph.Chart.DataGrid.SetData(i, 1, CDbl(Replace(frmMain.lstAdReport.ListItems.Item(i).ListSubItems.Item(ChartIndex).Text, "%", "")), 0)
    Next


    
    Chart.Title = ChartTitle
    Chart.ShowLegend = False
    
End Sub
Public Sub SetUpChartSearchReport(ChartTitle As String, xAxis As String, yAxis As String, ChartIndex As Integer)
    If frmMain.lstSearch.ListItems.Count < 1 Then
        MsgBox "No data to create chart", vbInformation
        Exit Sub
    End If

    frmGraph.Chart.Tag = "searchreport"
    frmGraph.Chart.Title.Text = ChartTitle
    frmGraph.Chart.Plot.Axis(VtChAxisIdX).AxisTitle.Text = xAxis
    frmGraph.Chart.Plot.Axis(VtChAxisIdY).AxisTitle.Text = yAxis
    
    Dim i As Long
    
    frmGraph.Chart.Plot.Wall.Brush.Style = VtBrushStyleSolid
    frmGraph.Chart.Plot.Wall.Brush.FillColor.Set 255, 255, 225
   

    frmGraph.Chart.RowCount = frmMain.lstSearch.ListItems.Count - 2
    frmGraph.Chart.ColumnCount = 1
 
    For i = 1 To frmGraph.Chart.RowCount
        
        frmGraph.Chart.Row = i
        frmGraph.Chart.RowLabel = frmMain.lstSearch.ListItems.Item(i).Text

       Call frmGraph.Chart.DataGrid.SetData(i, 1, CDbl(Replace(frmMain.lstSearch.ListItems.Item(i).ListSubItems.Item(ChartIndex).Text, "%", "")), 0)
    Next


    
    Chart.Title = ChartTitle
    Chart.ShowLegend = False
    
End Sub

Public Sub SetUpChartPaymentReport(ChartTitle As String, xAxis As String, yAxis As String, ChartIndex As Integer)
    If frmMain.lstPayment.ListItems.Count < 1 Then
        MsgBox "No data to create chart", vbInformation
        Exit Sub
    End If

    frmGraph.Chart.Tag = "payreport"
    frmGraph.Chart.Title.Text = ChartTitle
    frmGraph.Chart.Plot.Axis(VtChAxisIdX).AxisTitle.Text = xAxis
    frmGraph.Chart.Plot.Axis(VtChAxisIdY).AxisTitle.Text = yAxis
    Dim i As Long, Count As Long
    Dim cLabel() As String
    Dim cAmount() As Double
    ReDim cLabel(0)
    ReDim cAmount(0)
    Dim f As Boolean
    For i = 1 To frmMain.lstPayment.ListItems.Count
        f = False
        If CDbl(Replace(frmMain.lstPayment.ListItems.Item(i).ListSubItems.Item(ChartIndex).Text, "%", "")) > 0 Then

            cLabel(UBound(cLabel)) = frmMain.lstPayment.ListItems.Item(i).Text
            cAmount(UBound(cAmount)) = CDbl(Replace(frmMain.lstPayment.ListItems.Item(i).ListSubItems.Item(ChartIndex).Text, "%", ""))
            ReDim Preserve cLabel(UBound(cLabel) + 1)
            ReDim Preserve cAmount(UBound(cAmount) + 1)
            f = True
        End If
    
    Next
 
    If f = False Then
        ReDim cLabel(UBound(cLabel) - 1)
    End If
    
    frmGraph.Chart.Plot.Wall.Brush.Style = VtBrushStyleSolid
    frmGraph.Chart.Plot.Wall.Brush.FillColor.Set 255, 255, 225
   

    frmGraph.Chart.RowCount = UBound(cLabel)
    frmGraph.Chart.ColumnCount = 1
 
    For i = 1 To frmGraph.Chart.RowCount
        
        frmGraph.Chart.Row = i

        frmGraph.Chart.RowLabel = cLabel(i - 1)
        
       Call frmGraph.Chart.DataGrid.SetData(i, 1, cAmount(i - 1), 0)
    Next


    
    Chart.Title = ChartTitle
    Chart.ShowLegend = False
    
End Sub
Private Sub cmdSave_Click()
    #If Trial = True Then
        MsgBox "Not in Trial Version", vbExclamation
    #Else
        CD.FileName = ""
        CD.Flags = cdlOFNOverwritePrompt
        CD.DialogTitle = "Save Graph"
        CD.DefaultExt = ".bmp"
        CD.Filter = "BMP Files (*.bmp)|*.bmp"
        CD.ShowSave
        If CD.FileName <> "" Then
            Clipboard.Clear
            Chart.EditCopy
            Call SavePicture(Clipboard.GetData(vbCFDIB), CD.FileName)
        End If
    #End If
End Sub


Private Sub PrintChart()
    Printer.PaintPicture Clipboard.GetData(vbCFDIB), 0, 0
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Me.Tag = "2dBar"


    Chart.chartType = VtChChartType2dBar

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Chart.Width = Me.Width
    Chart.Height = Me.Height - 2200
    cmdPrint.Top = Me.Height - 1500
    cmdSave.Top = Me.Height - 1500
    cmdOptions.Top = Me.Height - 1500
    cboReport.Top = cmdOptions.Top - 500
    lblReportOptions.Top = cboReport.Top
    lblNote.Top = cmdOptions.Top - 500
End Sub

