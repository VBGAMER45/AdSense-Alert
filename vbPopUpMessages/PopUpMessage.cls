VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopUpMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4

Private WithEvents mfrmPopup           As frmPopUpMessage
Attribute mfrmPopup.VB_VarHelpID = -1

Private mstrKey             As String
Private mstrCaption         As String
Private mstrMessage         As String
Private mblnClickable       As Boolean
Private mlngForeColor       As OLE_COLOR
Private mlngBackColor       As OLE_COLOR
Private mpicLogo            As StdPicture
Private mpicBackground      As StdPicture
Private mstrWavFile         As String
Private mblnProgressBar     As Boolean
Private mlngValue           As Long
Private mblnSticky          As Boolean

Private mlngPercentShown    As Long
Private mblnFullyShown      As Boolean
Private mlngShowTime        As Long
Private mblnClosing         As Boolean

Private mlngParentPtr       As Long

Private Sub Class_Initialize()
    mlngBackColor = vbButtonFace
    mlngForeColor = vbBlack
End Sub

Private Sub Class_Terminate()
    If Not (mfrmPopup Is Nothing) Then
        Unload mfrmPopup
    End If
End Sub

Friend Sub Initialise(ByRef pobjParent As PopUpMessages)
Dim lngHeight       As Long
Dim lngWidth        As Long
    mblnFullyShown = False
    mblnClosing = False
    mlngPercentShown = 0
    mlngParentPtr = ObjPtr(pobjParent)
    Set mfrmPopup = New frmPopUpMessage
    Load mfrmPopup
    With mfrmPopup
        If pobjParent.ParentHandle <> 0 Then
            SetParent mfrmPopup.hWnd, pobjParent.ParentHandle
        End If
        .PopUpDirection = pobjParent.PopUpDirection
        .Width = pobjParent.MessageWidth
        .Height = pobjParent.MessageHeight
        lngHeight = .ScaleHeight
        lngWidth = .ScaleWidth
        Select Case pobjParent.PopUpDirection
            Case vbPopUpDirection.vbPopUp
                .Height = 0
                .Left = pobjParent.XPos
                .Top = pobjParent.YPos - .Height
            Case vbPopUpDirection.vbPopDown
                .Height = 0
                .Left = pobjParent.XPos
                .Top = pobjParent.YPos
            Case vbPopUpDirection.vbPopLeft
                .Width = 0
                .Left = pobjParent.XPos
                .Top = pobjParent.YPos
            Case vbPopUpDirection.vbPopRight
                .Width = 0
                .Left = pobjParent.XPos
                .Top = pobjParent.YPos
        End Select
        
        .lblCaption.Caption = mstrCaption
        .lblCaption.ForeColor = mlngForeColor
        .lblMessage.Caption = mstrMessage
        .lblMessage.FontUnderline = mblnClickable
        .lblMessage.MousePointer = IIf(mblnClickable, vbCustom, vbDefault)
        .lblMessage.ForeColor = mlngForeColor
        .BackColor = mlngBackColor
        Set .imgLogo.Picture = mpicLogo
        Set .imgBackground.Picture = mpicBackground
        .prgBar.Visible = mblnProgressBar
        .prgBar.Tag = mblnProgressBar
        .ResizeControls lngHeight, lngWidth
    End With
End Sub

Friend Property Get hWnd() As Long
    If Not (mfrmPopup Is Nothing) Then
        hWnd = mfrmPopup.hWnd
    End If
End Property

Friend Property Let ParenthWnd(ByVal Value As Long)
    SetParent mfrmPopup.hWnd, 0
End Property

Friend Property Get Parent() As PopUpMessages
    Set Parent = PtrObj(mlngParentPtr)
End Property

Public Property Let Key(ByVal Value As String)
    mstrKey = Value
End Property

Public Property Get Key() As String
    Key = mstrKey
End Property

Public Property Let Caption(ByVal Value As String)
    mstrCaption = Value
End Property

Public Property Get Caption() As String
    Caption = mstrCaption
End Property

Public Property Let Message(ByVal Value As String)
    mstrMessage = Value
End Property

Public Property Get Message() As String
    Message = mstrMessage
End Property

Public Property Let Clickable(ByVal Value As Boolean)
    mblnClickable = Value
End Property

Public Property Get Clickable() As Boolean
    Clickable = mblnClickable
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    mlngForeColor = Value
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mlngForeColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mlngBackColor = Value
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mlngBackColor
End Property

Public Property Set Logo(ByRef Value As StdPicture)
    Set mpicLogo = Value
End Property

Public Property Get Logo() As StdPicture
    Set Logo = mpicLogo
End Property

Public Property Set Background(ByRef Value As StdPicture)
    Set mpicBackground = Value
End Property

Public Property Get Background() As StdPicture
    Set Background = mpicBackground
End Property

Public Property Let WavFile(ByVal Value As String)
    mstrWavFile = Value
End Property

Public Property Get WavFile() As String
    WavFile = mstrWavFile
End Property

Public Property Let ProgressBar(ByVal Value As Boolean)
    mblnProgressBar = Value
End Property

Public Property Get ProgressBar() As Boolean
    ProgressBar = mblnProgressBar
End Property

Public Property Let Value(ByVal Value As Long)
    mlngValue = Value
    If Visible Then
        mfrmPopup.prgBar.Value = Value
    End If
End Property

Public Property Get Value() As Long
    Value = mlngValue
End Property

Public Property Let Sticky(ByVal Value As Boolean)
    mblnSticky = Value
End Property

Public Property Get Sticky() As Boolean
    Sticky = mblnSticky
End Property

Friend Property Let PercentShown(ByVal Value As Long)
    mlngPercentShown = Value
End Property

Friend Property Get PercentShown() As Long
    PercentShown = mlngPercentShown
End Property

Friend Property Let ShowTime(ByVal Value As Long)
    mlngShowTime = Value
End Property

Friend Property Get ShowTime() As Long
    ShowTime = mlngShowTime
End Property

Friend Property Let FullyShown(ByVal Value As Boolean)
    mblnFullyShown = True
End Property

Friend Property Get FullyShown() As Boolean
    FullyShown = mblnFullyShown
End Property

Public Property Get Visible() As Boolean
    Visible = Not (mfrmPopup Is Nothing)
End Property

Friend Property Get Finished() As Boolean
    Finished = ((mblnFullyShown Or mblnClosing) And mlngPercentShown = 0)
End Property

Friend Property Let Closing(ByVal Value As Boolean)
    mblnClosing = Value
End Property

Friend Property Get Closing() As Boolean
    Closing = mblnClosing
End Property

Friend Sub MoveMessage(ByVal plngLeft As Long, ByVal plngTop As Long, ByVal plngWidth As Long, ByVal plngHeight As Long)
    With mfrmPopup
        If Not .Visible Then
            .AlignHolder
            ShowWindow .hWnd, SW_SHOWNOACTIVATE
            SetWindowToTop .hWnd
            .Refresh
        End If
        .Move plngLeft, plngTop, plngWidth, plngHeight
    End With
End Sub

Friend Sub SetNewTranslucency()
    SetTranslucency mfrmPopup.hWnd, 255 * (PercentShown / 100)
End Sub

Public Sub ClosePopUp()
    mblnClosing = True
End Sub

Friend Sub DestroyPopup()
    Unload mfrmPopup
End Sub

Private Sub mfrmPopup_Click()
    If mblnClickable Then
        Parent.PopupClicked Me
    End If
    ClosePopUp
End Sub

Private Sub mfrmPopup_CloseGraceFully()
    ClosePopUp
End Sub

Private Sub mfrmPopup_Closing()
    mfrmPopup.Hide
    mfrmPopup.Refresh
    Parent.Remove Me
    Set mfrmPopup = Nothing
End Sub
