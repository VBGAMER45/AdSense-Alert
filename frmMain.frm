VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Adsense Alert - AdsenseAlert.com"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   11220
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   7185
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Version: "
            TextSave        =   "Version: "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Text            =   "Today's Clicks: "
            TextSave        =   "Today's Clicks: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4305
            MinWidth        =   4305
            Text            =   "Today's Impressions: "
            TextSave        =   "Today's Impressions: "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3952
            MinWidth        =   3952
            Text            =   "Todays Earnings: "
            TextSave        =   "Todays Earnings: "
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6975
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Login Information"
      TabPicture(0)   =   "frmMain.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgLock"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUsername"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPassword"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblChecking"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblBegin"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblBrowser"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtUsername"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdLogin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPassword"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCheckUpdate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdOptions"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdLock"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdHelp"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkSaveInformation"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "tmrAlertLoop"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdPurchase"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "tmrStatsLoop"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "frameTip"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Reports Ad Performance"
      TabPicture(1)   =   "frmMain.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdUpdateInterval"
      Tab(1).Control(1)=   "cmdViewCharts"
      Tab(1).Control(2)=   "cmdAddAlert"
      Tab(1).Control(3)=   "cmdPrintAdReports"
      Tab(1).Control(4)=   "cmdExportCSV"
      Tab(1).Control(5)=   "lstAdReport"
      Tab(1).Control(6)=   "frameShow"
      Tab(1).Control(7)=   "frameDateRange"
      Tab(1).Control(8)=   "Image1"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Search Performance"
      TabPicture(2)   =   "frmMain.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdViewSearchChart"
      Tab(2).Control(1)=   "cmdPrintSearch"
      Tab(2).Control(2)=   "cmdExportSearchCSV"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "lstSearch"
      Tab(2).Control(5)=   "Frame1"
      Tab(2).Control(6)=   "imgSearch"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Payments"
      TabPicture(3)   =   "frmMain.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdDisplayPayReport"
      Tab(3).Control(1)=   "cmdPayGraphs"
      Tab(3).Control(2)=   "cboYearPay2"
      Tab(3).Control(3)=   "cboMonthPay2"
      Tab(3).Control(4)=   "cboYearPay1"
      Tab(3).Control(5)=   "cboMonthPay1"
      Tab(3).Control(6)=   "cmdPrintPay"
      Tab(3).Control(7)=   "cmdSavePayments"
      Tab(3).Control(8)=   "lstPayment"
      Tab(3).Control(9)=   "Line2"
      Tab(3).Control(10)=   "lblDateRange"
      Tab(3).Control(11)=   "lblPaymentHistory"
      Tab(3).Control(12)=   "imgMoney"
      Tab(3).ControlCount=   13
      Begin VB.CommandButton cmdViewSearchChart 
         Caption         =   "View Charts"
         Height          =   495
         Left            =   -68520
         TabIndex        =   74
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrintSearch 
         Caption         =   "Print"
         Height          =   495
         Left            =   -67080
         TabIndex        =   73
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportSearchCSV 
         Caption         =   "Export CSV"
         Height          =   495
         Left            =   -65640
         TabIndex        =   72
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdateInterval 
         Caption         =   "Update Interval"
         Height          =   495
         Left            =   -74760
         TabIndex        =   53
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Frame frameTip 
         Caption         =   "Tip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   51
         Top             =   5400
         Width           =   5655
         Begin VB.Image imgTaskBar 
            Height          =   330
            Left            =   4320
            Picture         =   "frmMain.frx":68C2
            Top             =   480
            Width           =   1170
         End
         Begin VB.Label Label1 
            Caption         =   "Adsene Alert is also located in the system tray for easy access to current stats when you are logged in."
            Height          =   615
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Timer tmrStatsLoop 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   960
         Top             =   3360
      End
      Begin VB.CommandButton cmdPurchase 
         Caption         =   "Purchase Full Version"
         Height          =   1335
         Left            =   7440
         Picture         =   "frmMain.frx":7D4C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Timer tmrAlertLoop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   960
         Top             =   2760
      End
      Begin VB.CommandButton cmdDisplayPayReport 
         Caption         =   "Display Report"
         Height          =   495
         Left            =   -71040
         TabIndex        =   49
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkSaveInformation 
         Caption         =   "Save Login Information"
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         ToolTipText     =   "Saves login information if you have this program to autostart on restart"
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Show"
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
         Left            =   -67440
         TabIndex        =   47
         Top             =   360
         Width           =   3255
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   3015
            TabIndex        =   58
            Top             =   240
            Width           =   3015
            Begin VB.CommandButton cmdSearchSelectChannel 
               Caption         =   "Select Channels"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1800
               Picture         =   "frmMain.frx":E59E
               TabIndex        =   61
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optSearchChannelData 
               Caption         =   "Channel data "
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   60
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton optSearchAggregatedata 
               Caption         =   "Aggregate data "
               Height          =   255
               Left            =   0
               TabIndex        =   59
               Top             =   0
               Value           =   -1  'True
               Width           =   1455
            End
         End
      End
      Begin MSComctlLib.ListView lstSearch 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   46
         Top             =   1680
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Page Impressions"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clicks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Page CTR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Page eCPM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Your earnings "
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Date Range"
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
         Left            =   -73920
         TabIndex        =   38
         Top             =   360
         Width           =   6375
         Begin VB.ComboBox cboSearchPreset 
            Height          =   315
            ItemData        =   "frmMain.frx":14DF0
            Left            =   480
            List            =   "frmMain.frx":14E12
            TabIndex        =   45
            Text            =   "today"
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cboSearchMonth1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   480
            TabIndex        =   44
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboSearchDay1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   43
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboSearchYear1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2280
            TabIndex        =   42
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboSearchMonth2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            TabIndex        =   41
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboSearchDay2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   40
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboSearchYear2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            TabIndex        =   39
            Top             =   720
            Width           =   855
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   375
            TabIndex        =   67
            Top             =   240
            Width           =   375
            Begin VB.OptionButton optSearchRange 
               Height          =   255
               Left            =   0
               TabIndex        =   69
               Top             =   480
               Width           =   255
            End
            Begin VB.OptionButton optSearchPreset 
               Height          =   255
               Left            =   0
               TabIndex        =   68
               Top             =   0
               Value           =   -1  'True
               Width           =   255
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            ScaleHeight     =   615
            ScaleWidth      =   2415
            TabIndex        =   70
            Top             =   120
            Width           =   2415
            Begin VB.CommandButton cmdDisplalySearch 
               Caption         =   "Display Report"
               Height          =   495
               Left            =   0
               TabIndex        =   71
               Top             =   0
               Width           =   2055
            End
         End
         Begin VB.Line Line3 
            X1              =   3240
            X2              =   3480
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.CommandButton cmdPayGraphs 
         Caption         =   "Graphs"
         Height          =   1095
         Left            =   -68880
         Picture         =   "frmMain.frx":14EAA
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5040
         Width           =   1575
      End
      Begin VB.ComboBox cboYearPay2 
         Height          =   315
         Left            =   -65040
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboMonthPay2 
         Height          =   315
         Left            =   -65880
         TabIndex        =   35
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cboYearPay1 
         Height          =   315
         Left            =   -67080
         TabIndex        =   34
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboMonthPay1 
         Height          =   315
         Left            =   -67920
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdPrintPay 
         Caption         =   "Print"
         Height          =   1095
         Left            =   -67200
         Picture         =   "frmMain.frx":1B6FC
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdSavePayments 
         Caption         =   "Save"
         Height          =   1095
         Left            =   -65640
         Picture         =   "frmMain.frx":21F4E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5040
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstPayment 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   29
         Top             =   1320
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdViewCharts 
         Caption         =   "View Charts"
         Height          =   495
         Left            =   -68520
         TabIndex        =   27
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddAlert 
         Caption         =   "Add Alert"
         Height          =   495
         Left            =   -70200
         TabIndex        =   26
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrintAdReports 
         Caption         =   "Print"
         Height          =   495
         Left            =   -67080
         TabIndex        =   25
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportCSV 
         Caption         =   "Export CSV"
         Height          =   495
         Left            =   -65640
         TabIndex        =   24
         Top             =   6360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstAdReport 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   22
         Top             =   1680
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Page Impressions"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clicks"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Page CTR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Page eCPM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Your earnings "
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame frameShow 
         Caption         =   "Show"
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
         Left            =   -67440
         TabIndex        =   21
         Top             =   360
         Width           =   3255
         Begin VB.PictureBox picHold1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   3015
            TabIndex        =   54
            Top             =   240
            Width           =   3015
            Begin VB.CommandButton cmdSelectChannels 
               Caption         =   "Select Channels"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1800
               Picture         =   "frmMain.frx":287A0
               TabIndex        =   57
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optChannelData 
               Caption         =   "Channel data "
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   56
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton optAggregatedata 
               Caption         =   "Aggregate data "
               Height          =   255
               Left            =   0
               TabIndex        =   55
               Top             =   0
               Value           =   -1  'True
               Width           =   1455
            End
         End
      End
      Begin VB.Frame frameDateRange 
         Caption         =   "Date Range"
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
         Left            =   -73920
         TabIndex        =   13
         Top             =   360
         Width           =   6375
         Begin VB.ComboBox cboYear2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            TabIndex        =   20
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboDay2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   19
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboMonth2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            TabIndex        =   18
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboYear1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2280
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboDay1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cboMonth1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   480
            TabIndex        =   15
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboPresetDateRange 
            Height          =   315
            ItemData        =   "frmMain.frx":2EFF2
            Left            =   480
            List            =   "frmMain.frx":2F014
            TabIndex        =   14
            Text            =   "today"
            Top             =   240
            Width           =   2055
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   495
            TabIndex        =   62
            Top             =   240
            Width           =   495
            Begin VB.OptionButton optDateRange 
               Height          =   255
               Left            =   0
               TabIndex        =   64
               Top             =   480
               Width           =   375
            End
            Begin VB.OptionButton optDatePreset 
               Height          =   255
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Value           =   -1  'True
               Width           =   375
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   65
            Top             =   120
            Width           =   2295
            Begin VB.CommandButton cmdDisplayReport 
               Caption         =   "Display Report"
               Height          =   495
               Left            =   120
               TabIndex        =   66
               Top             =   0
               Width           =   2055
            End
         End
         Begin VB.Line Line1 
            X1              =   3240
            X2              =   3480
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   1095
         Left            =   6480
         Picture         =   "frmMain.frx":2F0AC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Need Help then click on this button."
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Set Password AdSense Alert"
         Height          =   1095
         Left            =   8760
         Picture         =   "frmMain.frx":358FE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Lock this program from others."
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   1095
         Left            =   6480
         Picture         =   "frmMain.frx":3C150
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Here you can edit the basic options of the program such as running on startup"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdCheckUpdate 
         Caption         =   "Check for Update"
         Height          =   1095
         Left            =   8760
         Picture         =   "frmMain.frx":429A2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Check if there are any updates to Adsense Alert"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Your adsense password"
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Default         =   -1  'True
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         ToolTipText     =   "Your adsense account email address"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   -74760
         Picture         =   "frmMain.frx":491F4
         Top             =   480
         Width           =   720
      End
      Begin VB.Image imgSearch 
         Height          =   720
         Left            =   -74760
         Picture         =   "frmMain.frx":4FA46
         Top             =   480
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   -66120
         X2              =   -65970
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblDateRange 
         Caption         =   "Date range:"
         Height          =   255
         Left            =   -68880
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPaymentHistory 
         Caption         =   "Payment History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73680
         TabIndex        =   28
         Top             =   480
         Width           =   5055
      End
      Begin VB.Image imgMoney 
         Height          =   720
         Left            =   -74760
         Picture         =   "frmMain.frx":56298
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblBrowser 
         Caption         =   "Launch Browser Window to Adsense Account"
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
         Left            =   1920
         TabIndex        =   23
         Top             =   4440
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label lblBegin 
         Caption         =   "To begin enter your adsene account and click on the login button"
         Height          =   495
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label lblChecking 
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   3840
         Width           =   4215
      End
      Begin VB.Label lblPassword 
         Caption         =   "Adsense Password:"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblUsername 
         Caption         =   "Adsense Username:"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Image imgLock 
         Height          =   720
         Left            =   1200
         Picture         =   "frmMain.frx":5CAEA
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   840
      End
   End
   Begin VB.Image imgBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":6333C
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileMinimizeToTray 
         Caption         =   "Minimize to Icon Tray"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAlerts 
      Caption         =   "&Alerts"
      Begin VB.Menu mnuAlertsAddAlert 
         Caption         =   "Add Alert"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAlertEditAlerts 
         Caption         =   "Edit Alerts"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAlertsViewAlerts 
         Caption         =   "Remove Alerts"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpCheckforUpdates 
         Caption         =   "&Check for updates"
      End
      Begin VB.Menu mnuHelpWebsite 
         Caption         =   "Visit &Website"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'AdSense Alert
'VisualBasicZone.com 2005-2006
'Jonathan Valentin
'********************************************
Option Explicit
Const Trial = False

Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
'MSN Messenger Popups style
''###JV Private WithEvents mobjPopups   As PopUpMessages
''###JV Private mobjDefault             As PopUpMessage

Private Sub cmdAddAlert_Click()
    frmAddAlert.Show
End Sub

Private Sub cmdCheckUpdate_Click()
    frmCheckUpdate.Show
End Sub

Private Sub cmdDisplalySearch_Click()
    If TrialExpired = True Then
        Exit Sub
    End If
    If optSearchPreset.Value = True Then
        If cboSearchPreset.Text = "today" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, , , True)
        ElseIf cboSearchPreset.Text = "yesterday" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "yesterday", , True)
        ElseIf cboSearchPreset.Text = "2 days ago" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "twodaysago", , True)
        ElseIf cboSearchPreset.Text = "last seven days" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "last7days", , True)
        ElseIf cboSearchPreset.Text = "this month" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "thismonth", , True)
        ElseIf cboSearchPreset.Text = "last month" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "lastmonth", , True)
        ElseIf cboSearchPreset.Text = "this week (Mon-Sun)" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "thisweek", , True)
        ElseIf cboSearchPreset.Text = "last week (Mon-Sun)" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "lastweek", , True)
        ElseIf cboSearchPreset.Text = "last business (Mon-Fri)" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "lastbusinessweek", , True)
        ElseIf cboSearchPreset.Text = "all time" Then
            Call modAdsense.GetTodaysAdData(Me.lstSearch, "alltime", , True)
        End If
    Else
    
    End If
End Sub

Private Sub cmdDisplayPayReport_Click()
    If TrialExpired = True Then
        Exit Sub
    End If
    Dim Month1 As Integer
    Month1 = cboMonthPay1.ListIndex + 1
    If Month1 = 0 Then Month1 = 1
    Dim month2 As Integer
    month2 = cboMonthPay2.ListIndex + 1
    If month2 = 0 Then month2 = 1
    Call modAdsense.GetPayment(lstPayment, Month1, month2, Me.cboYearPay1.Text, cboYearPay2.Text)
End Sub

Private Sub cmdDisplayReport_Click()
    If TrialExpired = True Then
        Exit Sub
    End If
    Dim strChannel As String
    strChannel = ""
    Dim i As Long
    If Me.optChannelData.Value = True Then
        
        For i = 0 To UBound(ChannelList)
            If ChannelList(i).Selected = True Then
                strChannel = strChannel & "&c.id=" & ChannelList(i).ChannelId
            End If
        Next i
        Debug.Print strChannel
    End If
    If optDatePreset.Value = True Then
        If cboPresetDateRange.Text = "today" Then
            If strChannel = "" Then
                Call modAdsense.GetTodaysAdData(Me.lstAdReport)
            Else
                Call modAdsense.GetTodaysChannelData(Me.lstAdReport, "today", strChannel)
            End If
            
        ElseIf cboPresetDateRange.Text = "yesterday" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "yesterday")
        ElseIf cboPresetDateRange.Text = "2 days ago" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "twodaysago")
        ElseIf cboPresetDateRange.Text = "last seven days" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "last7days")
        ElseIf cboPresetDateRange.Text = "this month" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "thismonth")
        ElseIf cboPresetDateRange.Text = "last month" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "lastmonth")
        ElseIf cboPresetDateRange.Text = "this week (Mon-Sun)" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "thisweek")
        ElseIf cboPresetDateRange.Text = "last week (Mon-Sun)" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "lastweek")
        ElseIf cboPresetDateRange.Text = "last business (Mon-Fri)" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "lastbusinessweek")
        ElseIf cboPresetDateRange.Text = "all time" Then
            Call modAdsense.GetTodaysAdData(Me.lstAdReport, "alltime")
        End If
    Else
        Dim Month1 As Integer
        Month1 = cboMonth1.ListIndex + 1
        If Month1 = 0 Then Month1 = 1
        Dim month2 As Integer
        month2 = cboMonth2.ListIndex + 1
        If month2 = 0 Then month2 = 1
        
        Call modAdsense.GetDateRange(lstAdReport, Month1, cboDay1.Text, cboYear1.Text, month2, cboDay2.Text, cboYear2.Text)
        
    End If
    
End Sub

Private Sub cmdExportCSV_Click()
'On Error GoTo errHandle
    #If Trial = True Then
        MsgBox "Not in trial version...", vbInformation
    #Else
        CD1.FileName = ""
        CD1.CancelError = False
        CD1.DialogTitle = "Export Ad Report CSV"
        CD1.DefaultExt = ".csv"
        CD1.Filter = "CSV Files (*.csv)|*.csv"
        CD1.Flags = cdlOFNOverwritePrompt
        CD1.ShowSave
        If CD1.FileName <> "" Then
            Dim f As Long
            f = FreeFile
            Open CD1.FileName For Output As #f
                Print #f, "Date, Page Impressions, Clicks, Page CTR, Page eCPM, Your earnings"
                Dim i As Long
                If Me.lstAdReport.ListItems.Count > 1 Then
                    For i = 1 To Me.lstAdReport.ListItems.Count
                        Print #f, Chr$(34) & Me.lstAdReport.ListItems.Item(i).Text & Chr$(34) & "," & lstAdReport.ListItems.Item(i).ListSubItems(1).Text & "," & lstAdReport.ListItems.Item(i).ListSubItems(2).Text & "," & lstAdReport.ListItems.Item(i).ListSubItems(3).Text & "," & lstAdReport.ListItems.Item(i).ListSubItems(4).Text & "," & lstAdReport.ListItems.Item(i).ListSubItems(5).Text
                    Next
                End If
                
                
            Close #f
            
        End If
    #End If
Exit Sub
errHandle:
    Call modGlobals.AddToErrorLog("Error_frmMain_cmdExportCSV: " & Err.Description & " " & Date & " " & Time)
End Sub

Private Sub cmdExportSearchCSV_Click()
On Error GoTo errHandle
    #If Trial = True Then
        MsgBox "Not in trial version...", vbInformation
    #Else
        CD1.FileName = ""
        CD1.CancelError = False
        CD1.DialogTitle = "Export Search Report CSV"
        CD1.DefaultExt = ".csv"
        CD1.Filter = "CSV Files (*.csv)|*.csv"
        CD1.Flags = cdlOFNOverwritePrompt
        CD1.ShowSave
        If CD1.FileName <> "" Then
            Dim f As Long
            f = FreeFile
            Open CD1.FileName For Output As #f
                Print #f, "Date, Page Impressions, Clicks, Page CTR, Page eCPM, Your earnings"
                
            Close #f
            
        End If
    #End If
Exit Sub
errHandle:
    Call modGlobals.AddToErrorLog("Error_frmMain_cmdExportSearchCSV_click: " & Err.Description & " " & Date & " " & Time)
    
End Sub

Private Sub cmdHelp_Click()
   ShellExecute Me.hWnd, vbNullString, App.Path & "\readme.txt", vbNullString, "C:\", SW_SHOWNORMAL

End Sub

Private Sub cmdLock_Click()
    frmSetPassword.Show vbModal, Me
End Sub

Private Sub cmdLogin_Click()
On Error GoTo errHandle
    Dim strData As String
    If TrialExpired = True Then
        Exit Sub
    End If
    
    If InternetCheckConnection("http://www.google.com", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        lblChecking.Caption = "Connection Failed! Check connection to internet!"
        lblChecking.ForeColor = vbRed
        lblChecking.FontBold = True
        Exit Sub
    Else
        lblChecking.ForeColor = vbGreen
        lblChecking.FontBold = False
    End If
    
    If txtUsername.Text <> "" And txtPassword.Text <> "" Then
       lblChecking.Caption = "Checking login information..."
       
            UserName = txtUsername.Text
            Password = txtPassword.Text
            UsernameSafe = Replace(UserName, "+", "%2B")
            PasswordSafe = Replace(Password, "+", "%2B")
      strData = GetPost("https://www.google.com/adsense/login.do", "username=" & UsernameSafe & "&password=" & PasswordSafe)
      
 
       Debug.Print strData
 

        bLoggedIn = False
        
        If InStr(strData, "Invalid email address or password") <> 0 Then
            MsgBox "Invaild Password!", vbCritical, "Wrong Password!"
            Exit Sub
        End If
        If InStr(strData, "Account Not Active") <> 0 Then
            MsgBox "That account doesn't exist!", vbExclamation, "No account!"
            Exit Sub
        End If
        'Search for username
      'Text1.Text = strData
        If InStr(strData, "Log Out</a>") <> 0 Then
            'Ok all ok
            bLoggedIn = True

            SSTab.TabVisible(1) = True
            SSTab.TabVisible(2) = True
            SSTab.TabVisible(3) = True
            Me.mnuAlertsAddAlert.Enabled = True
            Me.mnuAlertEditAlerts.Enabled = True
            Me.mnuAlertsViewAlerts.Enabled = True
            lblBrowser.Visible = True
            lblChecking.Caption = "You are logged in as " & txtUsername.Text
            'Get Channel List
            'https://www.google.com/adsense/report/aggregate?product=afc
            strData = GetPost("https://www.google.com/adsense/report/aggregate?product=afc", "username=" & UsernameSafe & "&password=" & PasswordSafe)
            Debug.Print strData
            Call modAdsense.GetUrlChannels(strData)
            
            'Get Todays Earnings
            Call modAdsense.GetTodaysAdData(Me.lstAdReport)
            tmrAlertLoop.Enabled = True
            tmrStatsLoop.Enabled = True
            
            Call UpdateStatusBar
          '  MsgBox "Couldn't Connect...", vbCritical
       End If
    Else
        MsgBox "You need to enter a username and a password!", vbInformation
        
    End If
Exit Sub
errHandle:
    Call AddToErrorLog("Error_frmMain_cmdLogin : " & Err.Description & " " & Date & " " & Time)
End Sub


Private Sub cmdOptions_Click()
    frmOptions.Show vbModal, Me
End Sub



Private Sub cmdPayGraphs_Click()
    If Me.lstPayment.ListItems.Count < 1 Then
        MsgBox "No data to create chart", vbInformation
        Exit Sub
    End If
    Unload frmGraph
    frmGraph.Chart.Tag = "payreport"
    frmGraph.Show
End Sub

Private Sub cmdPrintAdReports_Click()
    If Me.lstAdReport.ListItems.Count < 1 Then
        MsgBox "No data to print", vbInformation
        Exit Sub
    End If
    #If Trial = True Then
        MsgBox "Not in trial version...", vbInformation
    #Else
        On Error Resume Next
        CD1.CancelError = True

        CD1.DialogTitle = "Print Ad Reports"
        CD1.Flags = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection
        CD1.ShowPrinter
        Dim i As Long
        If Err <> cdlCancel Then
            
            For i = 1 To CD1.Copies
                Call modAdsense.PrintAdReport(lstAdReport, Me.cboPresetDateRange.Text)
            Next
        
        End If

    #End If
End Sub

Private Sub cmdPrintPay_Click()
    If Me.lstPayment.ListItems.Count < 1 Then
        MsgBox "No data to print", vbInformation
        Exit Sub
    End If
    
    #If Trial = True Then
        MsgBox "Not in trial version...", vbInformation
    #Else
        On Error Resume Next
        CD1.CancelError = True
        CD1.Flags = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection
        CD1.DialogTitle = "Print Payments"
        CD1.ShowPrinter
        Dim i As Long
        If Err <> cdlCancel Then
            For i = 1 To CD1.Copies
                Call modAdsense.PrintPayments(Me.lstPayment)
            Next
        End If
    #End If
End Sub

Private Sub cmdPrintSearch_Click()
On Error GoTo errHandle
    If Me.lstSearch.ListItems.Count < 1 Then
        MsgBox "No data to print", vbInformation
        Exit Sub
    End If
    #If Trial = True Then
        MsgBox "Not in trial version...", vbInformation
    #Else
        On Error Resume Next
        CD1.CancelError = True
        CD1.DialogTitle = "Print Search Reports"
        CD1.Flags = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection
        CD1.ShowPrinter
        Dim i As Long
        If Err <> cdlCancel Then
            For i = 1 To CD1.Copies
                
            Next
        End If
    #End If

Exit Sub
errHandle:
    Call modGlobals.AddToErrorLog("frmMain_cmdPrintSearch_click: " & Err.Description & " " & Date & " " & Time)

End Sub

Private Sub cmdPurchase_Click()
    On Error Resume Next
     ShellExecute Me.hWnd, vbNullString, "http://www.adsensealert.com/order.php", vbNullString, "C:\", SW_SHOWNORMAL
     
End Sub

Private Sub cmdSavePayments_Click()
On Error GoTo errHandle
    #If Trial = True Then
        MsgBox "Not in trial version...", vbInformation
    #Else
        CD1.FileName = ""
        CD1.CancelError = False
        CD1.DialogTitle = "Export Payments CSV"
        CD1.Flags = cdlOFNOverwritePrompt
        CD1.Filter = "CSV Files (*.csv)|*.csv"
        CD1.DefaultExt = ".csv"
        CD1.ShowSave
        If CD1.FileName <> "" Then
            Dim f As Long
            f = FreeFile
            Open CD1.FileName For Output As #f
    
                Print #f, "Date, Description, Amount"
                Dim i As Long
                If Me.lstPayment.ListItems.Count > 1 Then
                    For i = 1 To Me.lstPayment.ListItems.Count
                        Print #f, Chr$(34) & Me.lstPayment.ListItems.Item(i).Text & Chr$(34) & "," & Chr$(34) & lstPayment.ListItems.Item(i).ListSubItems(1).Text & Chr$(34) & "," & Chr$(34) & lstPayment.ListItems.Item(i).ListSubItems(2).Text & Chr$(34)
                    Next
                End If
                
                
            Close #f
            
        End If

    #End If
Exit Sub
errHandle:
    Call modGlobals.AddToErrorLog("frmMain_cmdSavePayments_click: " & Err.Description & " " & Date & " " & Time)
End Sub

Private Sub cmdSelectChannels_Click()
    frmSelectChannel.Show
End Sub

Private Sub cmdUpdateInterval_Click()
    frmOptions.Show
End Sub

Private Sub cmdViewCharts_Click()
 
    If Me.lstAdReport.ListItems.Count < 1 Then
        MsgBox "No data to create chart", vbInformation
        Exit Sub
    End If
    Unload frmGraph
    frmGraph.Chart.Tag = "adreport"
    frmGraph.Chart.Title.Text = "Adsense Alert Impressions Report"
    frmGraph.Chart.ColumnCount = lstAdReport.ListItems.Count - 2
    frmGraph.Chart.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "Date"
    frmGraph.Chart.Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Impressions"
    
    Dim arrData() As String
    ReDim arrData(lstAdReport.ListItems.Count - 2)
    Dim i As Long
    
    For i = 1 To UBound(arrData)
        
        
        arrData(i - 1) = lstAdReport.ListItems.Item(i).ListSubItems.Item(1).Text

       ' arrData(i - 1, 5) = lstAdReport.ListItems.Item(i).ListSubItems.Item(5).Text
        'arrData(0, i - 1) = Me.lstAdReport.ListItems.Item(i).Text
        'arrData(1, i - 1) = Me.lstAdReport.ListItems.Item(i).ListSubItems.Item(1).Text
        'arrData(2, i - 1) = Me.lstAdReport.ListItems.Item(i).ListSubItems.Item(2).Text
        'arrData(3, i - 1) = Me.lstAdReport.ListItems.Item(i).ListSubItems.Item(3).Text
        'arrData(4, i - 1) = Me.lstAdReport.ListItems.Item(i).ListSubItems.Item(4).Text
        'arrData(5, i - 1) = Me.lstAdReport.ListItems.Item(i).ListSubItems.Item(5).Text
    Next
    ReDim Preserve arrData(UBound(arrData) - 1)

    'frmGraph.Chart.ChartData = arrData
    frmGraph.Chart.Plot.Wall.Brush.Style = VtBrushStyleSolid
    frmGraph.Chart.Plot.Wall.Brush.FillColor.Set 255, 255, 225
   

    frmGraph.Chart.RowCount = UBound(arrData) + 1
    frmGraph.Chart.ColumnCount = UBound(arrData) + 1
 
    For i = 1 To frmGraph.Chart.RowCount
        
        frmGraph.Chart.Row = i
        frmGraph.Chart.RowLabel = lstAdReport.ListItems.Item(i).Text
        Dim k As Long
       
       Call frmGraph.Chart.DataGrid.SetData(i, 1, CDbl(lstAdReport.ListItems.Item(i).ListSubItems.Item(1).Text), 0)
    Next
    
    For i = 1 To UBound(arrData) + 1

    frmGraph.Chart.DataGrid.ColumnLabel(i, 1) = lstAdReport.ListItems.Item(i).Text
    'frmGraph.Chart.DataGrid.r(1, 1) = lstAdReport.ListItems.Item(i).Text

        With frmGraph.Chart.Plot.SeriesCollection(i).DataPoints(-1)
            .DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
            '.DataPointLabel.ValueFormat = "0.00"
            .DataPointLabel.VtFont.Name = "Tahoma"
            .DataPointLabel.VtFont.Size = 8
            .DataPointLabel.VtFont.Style = VtFontStyleOutline 'Regular
            '.DataPointLabel.VtFont.Style = VtFontStyleBold Or VtFontStyleItalic 'Do both Bold AND Italic
            
            'Label color by serie
           ' If i = 1 Then
           '     .DataPointLabel.VtFont.VtColor.Set 255, 64, 64
           ' ElseIf i = 2 Then
           '     .DataPointLabel.VtFont.VtColor.Set 128, 128, 255
           ' ElseIf i = 3 Then
                '.DataPointLabel.VtFont.VtColor.Set 64, 192, 192
            'End If
        
            'Show marker
             frmGraph.Chart.Plot.SeriesCollection(1).SeriesMarker.Show = False
        End With
Next

    frmGraph.Chart.chartType = VtChChartType2dBar
    frmGraph.Show
End Sub

Private Sub cmdViewSearchChart_Click()

    If Me.lstSearch.ListItems.Count < 1 Then
        MsgBox "No data to create chart", vbInformation
        Exit Sub
    End If
    Unload frmGraph
    frmGraph.Chart.Tag = "searchreport"
    frmGraph.Show
End Sub

Private Sub Form_Load()
    TrialExpired = False
    Me.Tag = "adsensealert.com visualbasiczone.com copyright Jonathan Valentin 2006"
    RunOnStartUp = True
    #If Trial = True Then
        If DateGood(3) = False Then
            TrialExpired = True
            frmExpired.Show
            Me.Hide
            Exit Sub
        End If
        
        Call CheckTrial
    #End If
    'Msn
   '###JV Set mobjPopups = New PopUpMessages
    UseMsnUpdates = False '###JVTrue
    SoundOnUpdate = True
    'Email settings
    EmailUpdates = False
    EmailUrl = "http://www.site.com/mail.php"
    'Setup combo boxes
    Call SetupComboDate(cboDay1, False, False, True)
    Call SetupComboDate(cboDay2, False, False, True)
    Call SetupComboDate(cboMonth1, True, False, False)
    Call SetupComboDate(cboMonth2, True, False, False)
    Call SetupComboDate(cboYear1, False, True, False)
    Call SetupComboDate(cboYear2, False, True, False)
    Call SetupComboDate(cboMonthPay1, True, False, False)
    Call SetupComboDate(cboMonthPay2, True, False, False)
    Call SetupComboDate(cboSearchMonth1, True, False, False)
    Call SetupComboDate(cboSearchMonth2, True, False, False)
    Call SetupComboDate(cboSearchDay1, False, False, True)
    Call SetupComboDate(cboSearchDay2, False, False, True)
    Call SetupComboDate(cboSearchYear1, False, True, False)
    Call SetupComboDate(cboSearchYear2, False, True, False)

    'Setup Payment year range
    Dim i As Integer
    For i = 0 To 7
        Dim k As Long
        k = 2003 + i
        cboYearPay1.Text = 2003
        cboYearPay2.Text = 2015
        cboYearPay1.AddItem k
        cboYearPay2.AddItem k
    Next
        
    StatusBar.Panels(1).Text = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    'StatusBar.Panels.Item(2).Visible = False
    'StatusBar.Panels.Item(3).Visible = False
    'StatusBar.Panels.Item(4).Visible = False
    StatusBar.Panels.Item(2).Visible = True
    StatusBar.Panels.Item(3).Visible = True
    StatusBar.Panels.Item(4).Visible = True
    
    #If Trial = True Then
        StatusBar.Panels(1).Width = 2200
        StatusBar.Panels(1).Text = "TRIAL VERSION: " & App.Major & "." & App.Minor & "." & App.Revision
        Me.Caption = "Adsense Alert - AdsenseAlert.com -  Trial Version"
        cmdPurchase.Visible = True
    #End If
    

    
    'Show the tray icon
    IsInTray = False
    frmTrayIcon.Show
    frmTrayIcon.Hide
    'Hide Tabs
    'SSTab.TabVisible(1) = False
    'SSTab.TabVisible(2) = False
    'SSTab.TabVisible(3) = False
    
    SSTab.TabVisible(1) = True
    SSTab.TabVisible(2) = True
    SSTab.TabVisible(3) = True

    'Load Config information
    Call LoadConfigInformation
    
    If UpdateInterval = 0 Then
        UpdateInterval = 10
    End If
    'Resize the alert list
    ReDim AlertList(0)


    'Used on restart
    If Command = "tray" Then
        'Login to account
        Me.Hide
        cmdLogin_Click
        Me.Hide
    End If
    
    #If Trial = True Then
        If DateGood(3) = False Then
            TrialExpired = True
            frmExpired.Show
            Me.Hide
            Exit Sub
        End If
    #End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    Me.Hide
    IsInTray = True
    frmTrayIcon.mnuHideAdsenseAlert.Caption = "Show Adsense Alert"
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
       '###JV Set mobjPopups = Nothing
       '###JV Set mobjDefault = Nothing
End Sub

Private Sub imgTaskBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgTaskBar.ToolTipText = modGlobals.MySysTray.TipText
    
End Sub

Private Sub lblBrowser_Click()
    On Error Resume Next
     ShellExecute Me.hWnd, vbNullString, "https://www.google.com/adsense/login.do?username=" & UsernameSafe & "&password=" & PasswordSafe, vbNullString, "C:\", SW_SHOWNORMAL
     
End Sub

Private Sub mnuAlertEditAlerts_Click()
    frmEditAlerts.Show
End Sub

Private Sub mnuAlertsAddAlert_Click()
    frmAddAlert.Show
End Sub

Private Sub mnuAlertsViewAlerts_Click()
    frmViewAlerts.Show
End Sub

Private Sub mnuFileExit_Click()
    Dim Response As VbMsgBoxResult
    Response = MsgBox("Are you sure you want to quit?", vbYesNo + vbInformation, "Quit Adsense Alert?")
    
    If Response = vbYes Then
        Call SaveConfigInformation
        Unload frmOptions
        Unload frmSetPassword
        Unload frmAbout
        Unload frmMain
        Unload Me
        End
    End If
End Sub

Private Sub mnuFileMinimizeToTray_Click()
    Me.Hide
    IsInTray = True
    frmTrayIcon.mnuHideAdsenseAlert.Caption = "Show Adsense Alert"
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpCheckforUpdates_Click()
    frmCheckUpdate.Show
    
End Sub

Private Sub mnuHelpWebsite_Click()
    On Error Resume Next
     ShellExecute Me.hWnd, vbNullString, "http://www.adsensealert.com", vbNullString, "C:\", SW_SHOWNORMAL
     
End Sub

Private Sub mnuOptions_Click()
On Error GoTo errHandle
    Unload frmOptions
    Unload frmAbout
    Unload frmViewAlerts
    Unload frmAlert
    Unload frmAddAlert
    frmOptions.Show vbModal, Me
Exit Sub
errHandle:

End Sub

Private Sub optAggregatedata_Click()
    If optAggregatedata.Value = False Then
        cmdSelectChannels.Enabled = True
    Else
        cmdSelectChannels.Enabled = False
    End If
End Sub

Private Sub optChannelData_Click()
    If optChannelData.Value = True Then
        cmdSelectChannels.Enabled = True
    Else
        cmdSelectChannels.Enabled = False
    End If
End Sub

Private Sub optDatePreset_Click()
    If optDateRange.Value = True Then
        cboPresetDateRange.Enabled = False
    Else
        cboPresetDateRange.Enabled = True
        cboMonth1.Enabled = False
        cboMonth2.Enabled = False
        cboYear1.Enabled = False
        cboYear2.Enabled = False
        cboDay1.Enabled = False
        cboDay2.Enabled = False
    End If
End Sub

Private Sub optDateRange_Click()
    If optDateRange.Value = True Then
        cboPresetDateRange.Enabled = False
        cboMonth1.Enabled = True
        cboMonth2.Enabled = True
        cboYear1.Enabled = True
        cboYear2.Enabled = True
        cboDay1.Enabled = True
        cboDay2.Enabled = True
        
    Else
        cboPresetDateRange.Enabled = True
    End If
End Sub

Private Sub optSearchAggregatedata_Click()
    If optSearchAggregatedata.Value = False Then
       cmdSearchSelectChannel.Enabled = True
    Else
        cmdSearchSelectChannel.Enabled = False
    End If
End Sub

Private Sub optSearchChannelData_Click()
    If optSearchChannelData.Value = True Then
       cmdSearchSelectChannel.Enabled = True
    Else
        cmdSearchSelectChannel.Enabled = False
    End If
End Sub

Private Sub optSearchPreset_Click()
    If optSearchPreset.Value = True Then
        cboSearchPreset.Enabled = True
        cboSearchMonth1.Enabled = False
        cboSearchMonth2.Enabled = False
        cboSearchDay1.Enabled = False
        cboSearchDay2.Enabled = False
        cboSearchYear1.Enabled = False
        cboSearchYear2.Enabled = False
        
    Else
        cboSearchPreset.Enabled = False
        
    End If
End Sub

Private Sub optSearchRange_Click()
    If optSearchRange.Value = True Then
        cboSearchPreset.Enabled = False
        cboSearchMonth1.Enabled = True
        cboSearchMonth2.Enabled = True
        cboSearchDay1.Enabled = True
        cboSearchDay2.Enabled = True
        cboSearchYear1.Enabled = True
        cboSearchYear2.Enabled = True
        
    Else
        cboSearchPreset.Enabled = False
    End If
End Sub

Sub SetupComboDate(cboBox As ComboBox, month As Boolean, year As Boolean, day As Boolean)
    Dim i As Long
    If day = True Then
        cboBox.Clear
        cboBox.Text = "1"
        For i = 1 To 31
            cboBox.AddItem i
        Next
    End If
    If year = True Then
        cboBox.Clear
        cboBox.Text = "2001"
        For i = 1 To 11
            Dim k As Long
            k = 2000 + i
            cboBox.AddItem k
        Next
    End If
    If month = True Then
        cboBox.Clear
        cboBox.Text = "Jan"
        cboBox.AddItem "Jan"
        cboBox.AddItem "Feb"
        cboBox.AddItem "Mar"
        cboBox.AddItem "Apr"
        cboBox.AddItem "May"
        cboBox.AddItem "Jun"
        cboBox.AddItem "Jul"
        cboBox.AddItem "Aug"
        cboBox.AddItem "Sep"
        cboBox.AddItem "Oct"
        cboBox.AddItem "Nov"
        cboBox.AddItem "Dec"
    End If
End Sub
Public Sub UpdateStatusBar()

    
    StatusBar.Panels.Item(2).Text = "Today's Clicks: " & TodayClicks
    StatusBar.Panels.Item(3).Text = "Today's Impressions: " & TodayImpressions
    StatusBar.Panels.Item(4).Text = "Today's Earnings: " & TodayEarnings
    StatusBar.Panels.Item(2).Visible = True
    StatusBar.Panels.Item(3).Visible = True
    StatusBar.Panels.Item(4).Visible = True
    StatusBar.Refresh
    
    MySysTray.PopUpMessage = "Clicks: " & TodayClicks & " Impressions: " & TodayImpressions & " Earnings: " & TodayEarnings
    MySysTray.TipText = MySysTray.PopUpMessage
End Sub
Public Sub SaveConfigInformation()
On Error GoTo errHandle
    Dim f As Long
    f = FreeFile
    Open App.Path & "\config.ini" For Output As #f
        Print #f, "Config File for AdsenseAlert"
        #If Trial = True Then
            Print #f, "Trial Version"
        #End If
        If chkSaveInformation.Value = vbChecked Then
            Print #f, "Username=" & UserName
            Print #f, "Password=" & modGlobals.Rot39(Password)
        End If
        If RunOnStartUp = True Then
            Print #f, "Startup"
        End If
        If EmailUpdates = True Then
            Print #f, "EmailUpdates"
        End If
        If SoundOnUpdate = True Then
            Print #f, "Sound"
        Else
            Print #f, "NoSound"
        End If
        If UseMsnUpdates = True Then
            Print #f, "MSNUpdate"
        Else
            Print #f, "NoMSNUpdate"
        End If
        
        Print #f, "Emailurl=" & EmailUrl
        
        Print #f, "Interval=" & UpdateInterval
        If AdsenseAlertPassword <> "" Then
            Print #f, "AlertPassword=" & modGlobals.Rot39(AdsenseAlertPassword)
        End If
        Print #f, "AlertList"
        Dim i As Long
        For i = 0 To UBound(AlertList)
            If AlertList(i).Amount <> 0 Then
                Print #f, AlertList(i).AlertType & "," & AlertList(i).ConditionType & "," & AlertList(i).Amount
            End If
        Next
    Close #f
Exit Sub
errHandle:
    MsgBox "Error_frmMain_SaveConfigInformation: " & Err.Description
End Sub
Public Sub LoadConfigInformation()
On Error GoTo errHandle
    Dim f As Long
    Dim strData As String
    f = FreeFile
    #If Trial = True Then
        Dim TFound As Boolean
        TFound = False
    #End If
    Open App.Path & "\config.ini" For Input As #f
        Do While Not EOF(f)
            Line Input #f, strData

            If Left$(strData, 9) = "Username=" Then
                 UserName = Right$(strData, Len(strData) - 9)
            End If
            If Left$(strData, 9) = "Password=" Then
                 Password = modGlobals.Rot39(Right$(strData, Len(strData) - 9))
            End If
            If Left$(strData, 9) = "Interval=" Then
                 UpdateInterval = Right$(strData, Len(strData) - 9)
            End If
            If Left$(strData, 9) = "Emailurl=" Then
                 EmailUrl = Right$(strData, Len(strData) - 9)
            End If
            If Left$(strData, 14) = "AlertPassword=" Then
                 AdsenseAlertPassword = modGlobals.Rot39(Right$(strData, Len(strData) - 14))
            End If
            If Left$(strData, 6) = "Startup" Then
                RunOnStartUp = True
            End If
            If Left$(strData, 12) = "EmailUpdates" Then
                EmailUpdates = False
            End If
            If Left$(strData, 5) = "Sound" Then
                SoundOnUpdate = True
            End If
            If Left$(strData, 7) = "NoSound" Then
                SoundOnUpdate = False
            End If
            If Left$(strData, 8) = "MSNUpdate" Then
                UseMsnUpdates = True
            End If
            If Left$(strData, 10) = "NoMSNUpdate" Then
                UseMsnUpdates = False
            End If
            
            #If Trial = True Then
                If Left$(strData, 13) = "Trial Version" Then
                     TFound = True
                End If
            #End If
        Loop
    Close #f
    #If Trial = True Then
    If TFound = False Then
        MsgBox "Nice try but thats not going to work. I suggest you stop trying hehe."
    End If
    #End If
    txtUsername.Text = UserName
    txtPassword.Text = Password

    If RunOnStartUp = True Then
        Call modGlobals.RegRun(App.Path & "\AdsenseAlert.exe tray", "AdsenseAlert")
    Else
        Call modGlobals.RemoveRegRun("AdsenseAlert")
    End If
    
    Exit Sub
errHandle:
    Exit Sub
    MsgBox "Error_frmMain_LoadConfigInformation: " & Err.Description
End Sub

Private Sub tmrAlertLoop_Timer()
    Dim i As Long
    
    
    For i = 0 To UBound(AlertList)
        If AlertList(i).AlertDate <> Date Then
            AlertList(i).AlertOn = True
        End If
        
        If AlertList(i).AlertOn = True Then
        'clicks, earnings, impressions
            If AlertList(i).ConditionType = 1 Then
                '>
                Select Case AlertList(i).AlertType
                
                    Case 1
                        If TodayClicks > AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's Clicks are greater than " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    Case 2
                        If TodayEarnings > AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's Earnings are greater than " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    Case 3
                        If TodayImpressions > AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's Impressions are greater than " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    Case 4
                        If TodayCTR > AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's CTR are greater than " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    
                End Select
            End If
            If AlertList(i).ConditionType = 2 Then
                '=
                Select Case AlertList(i).AlertType
                
                    Case 1
                        If TodayClicks = AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's Clicks have equaled " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    Case 2
                        If TodayEarnings = AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's Earnings have equaled " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    Case 3
                        If TodayImpressions = AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's Impressions have equaled " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                    Case 4
                        If TodayCTR = AlertList(i).Amount Then
                            Load frmAlert
                            frmAlert.lblMessage.Caption = "Today's CTR have equaled " & AlertList(i).Amount
                            frmAlert.Show
                            AlertList(i).AlertDate = Date
                            AlertList(i).AlertOn = False
                        End If
                End Select
            End If
        
        End If
    
    Next
    tmrAlertLoop.Enabled = False
End Sub

Private Sub tmrStatsLoop_Timer()
    CurrentUpdate = CurrentUpdate + 1
    'Update the stats
    If CurrentUpdate >= UpdateInterval Then
        Call modAdsense.GetTodaysAdData(Me.lstAdReport, , False)
        tmrAlertLoop.Enabled = True
        CurrentUpdate = 0
    End If
End Sub
Public Sub AddPopup(ByVal strTitle As String, ByVal strMessage As String, Optional IsSticky As Boolean = False)

Exit Sub '###JV
'###JVDim objPopUp    As PopUpMessage
 '###JVmobjPopups.AllowFading = True
   '###JV Set objPopUp = New PopUpMessage
    '###JVWith objPopUp
       '###JV .Caption = strTitle
       '###JV .Message = strMessage
      '###JV  .Clickable = False
        
        '###JV.Sticky = IsSticky
        
       '###JV Set .Background = imgBack.Item(0)


        '###JVSet .Logo = Me.icon

        If SoundOnUpdate = True Then
         '###JV   .WavFile = App.Path & "\new.wav"
        Else
         '###JV   .WavFile = ""
        End If
        
   '###JV End With
    
  '###JV  mobjPopups.Show objPopUp
End Sub
