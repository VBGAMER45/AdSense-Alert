Attribute VB_Name = "modGlobals"
'********************************************
'AdSense Alert
'VisualBasicZone.com 2005
'Jonathan Valentin
'********************************************
Option Explicit

'Current Version of Adsense Alert
'Public Const strVersion As String = "0.01"

#Const Trial = False

Public bLoggedIn As Boolean
Public IsRegistered As Boolean
Public IsInTray As Boolean
Public AdsenseAlertPassword As String
Public IsLocked As Boolean
Public RunOnStartUp As Boolean
Public UseMsnUpdates As Boolean
Public SoundOnUpdate As Boolean
Public EmailUpdates As Boolean
Public EmailAlerts As Boolean
Public EmailUrl As String

Public UserName As String
Public Password As String
Public UpdateInterval As Long
Public CurrentUpdate As Long

'Username and Password where + is handled
Public UsernameSafe As String
Public PasswordSafe As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'Todays Stats
Public TodayClicks As Long
Public TodayImpressions As Long
Public TodayEarnings As Single
Public TodayCTR As Single
Public TodayCPM As Single

'Channels
Private Type ChannelListType
    ChannelId As String
    ChannelName As String
    Selected As Boolean
End Type
Public ChannelList() As ChannelListType

Public MySysTray As New CSystrayIcon

'Alerts
Private Type AlertType
    AlertType As Byte
    ConditionType As Byte
    Amount As Long
    AlertDate As String
    AlertOn As Boolean
End Type

Public AlertList() As AlertType

'Windows XP Controls
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200


Private Const ECM_FIRST As Long = &H1500
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)

Private Type EDITBALLOONTIP
   cbStruct As Long
   pszTitle As String
   pszText As String
   ttiIcon As Long
End Type

Private Declare Function SendMessage Lib "User32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

'Rot 39 Encryption
Private Const LOWER_LIMIT As Long = 48   'ascii for 0
Private Const UPPER_LIMIT As Long = 125  'ascii for {
Private Const CHARMAP     As Long = 39
'For Trial
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public DaysLeft As Long
Public TrialExpired As Boolean
'For Reg
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3 ' Free form binary
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long




Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Save a string to the key
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey Ret
End Sub
Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Delete the key's value
    RegDeleteValue Ret, strValue
    'close the key
    RegCloseKey Ret
End Sub
Public Sub CheckTrial()
TrialExpired = False
Exit Sub
    Dim str1 As String, str2 As String
    Dim strKey As String
    'Adsense Alert
    'str1 = Chr$(65) & Chr$(100) & Chr$(115) & Chr$(101) & Chr$(110) & Chr$(115) & Chr$(101) & Chr$(32) & Chr$(65) & Chr$(108) & Chr$(101) & Chr$(114) & Chr$(116)
    'Options
    'str2 = Chr$(79) & Chr$(112) & Chr$(116) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(115)
    'Tray
    'strKey = Chr$(84) & Chr$(114) & Chr$(97) & Chr$(121)
    'Dim strdata As String
    '
   ' strdata = GetSetting(str1, "Options", strKey, "")
    
    
    If DateGood(3) = False Then
        TrialExpired = True
    End If
    
        Dim sSave As String, Ret As Long
        'Create a buffer
        sSave = Space(255)
        'Get the system directory
        Ret = GetSystemDirectory(sSave, 255)
        'Remove all unnecessary chr$(0)'s
        sSave = Left$(sSave, Ret)
        Dim f As Long
        f = FreeFile
        sSave = sSave & "\" & Chr$(97) & Chr$(100) & Chr$(115) & Chr$(117) & Chr$(112) & Chr$(100) & Chr$(97) & Chr$(116) & Chr$(101) & Chr$(46) & Chr$(116) & Chr$(120) & Chr$(116)
    If FileExists(sSave) = True Then
        TrialExpired = True
    End If
    
    If TrialExpired = True Then

        Open sSave For Output As #f
            Print #f, App.Major & "." & App.Minor & "." & App.Revision
        Close #f
        
        MySysTray.TipText = Chr$(84) & Chr$(114) & Chr$(105) & Chr$(97) & Chr$(108) & Chr$(32) & Chr$(69) & Chr$(120) & Chr$(112) & Chr$(105) & Chr$(114) & Chr$(101) & Chr$(100)
        frmExpired.Show
        frmMain.Hide
        Exit Sub
    End If
    
   
End Sub
Public Function FileExists(ByVal Path As String) As Boolean
'*****************************
'Purpose: Checks wether a FileExists or not
'*****************************
  If Len(Path) = 0 Then Exit Function
  If Dir(Path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> vbNullString Then FileExists = True
End Function
Public Function DateGood(NumDays As Integer) As Boolean
    'The purpose of this module is to allow you to place a time
    'limit on the unregistered use of your shareware application.
    'This module can not be defeated by rolling back the system clock.
    'Simply call the DateGood function when your application is first
    'loading, passing it the number of days it can be used without
    'registering.
    '
    'Ex: If DateGood(30)=False Then
    ' CrippleApplication
    ' End if
    'Register Parameters:
    ' CRD: Current Run Date
    ' LRD: Last Run Date
    ' FRD: First Run Date

    Dim TmpCRD As Date
    Dim TmpLRD As Date
    Dim TmpFRD As Date
    
    Dim str2 As String
    'Options
    str2 = Chr$(79) & Chr$(112) & Chr$(116) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(115)
    
    TmpCRD = Format(Now, "m/d/yy")
    TmpLRD = Rot39(GetSetting(App.EXEName, str2, "LRD", Rot39("1/1/2000")))
    TmpFRD = Rot39(GetSetting(App.EXEName, str2, "FRD", Rot39("1/1/2000")))
    DateGood = False

    'If this is the applications first load, write initial settings
    'to the register
    If TmpLRD = "1/1/2000" Then
        SaveSetting App.EXEName, str2, "LRD", Rot39(TmpCRD)
        SaveSetting App.EXEName, str2, "FRD", Rot39(TmpCRD)
    End If
    'Read LRD and FRD from register
    TmpLRD = Rot39(GetSetting(App.EXEName, str2, "LRD", Rot39("1/1/2000")))
    TmpFRD = Rot39(GetSetting(App.EXEName, str2, "FRD", Rot39("1/1/2000")))

    If TmpFRD > TmpCRD Then 'System clock rolled back
        DateGood = False
    ElseIf Now > DateAdd("d", NumDays, TmpFRD) Then 'Expiration expired
        DateGood = False
    ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
        SaveSetting App.EXEName, str2, "LRD", Rot39(TmpCRD)
        DateGood = True
    ElseIf TmpCRD = Format(TmpLRD, "m/d/yy") Then
        DateGood = True
    Else
        DateGood = False
    End If
End Function
Sub Main()
    Call InitCommonControlsVB
    
    #If Trial = True Then
        Dim sSave As String, Ret As Long
        'Create a buffer
        sSave = Space(255)
        'Get the system directory
        Ret = GetSystemDirectory(sSave, 255)
        'Remove all unnecessary chr$(0)'s
        sSave = Left$(sSave, Ret)
        
        sSave = sSave & "\" & Chr$(97) & Chr$(100) & Chr$(115) & Chr$(117) & Chr$(112) & Chr$(100) & Chr$(97) & Chr$(116) & Chr$(101) & Chr$(46) & Chr$(116) & Chr$(120) & Chr$(116)
        If DateGood(3) = False Then
            TrialExpired = True
            Dim f As Long
            f = FreeFile
            Open sSave For Output As #f
                Print #f, App.Major & "." & App.Minor & "." & App.Revision
            Close #f
            frmExpired.Show

            Exit Sub
        End If

  
        f = FreeFile
        
        If FileExists(sSave) = True Then
            TrialExpired = True
            frmExpired.Show

            Exit Sub
        End If
        
        Call modGlobals.CheckTrial
        
    #End If
    
    frmMain.Show
End Sub
Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Function URLEncode(ByVal urlText As String) As String

    Dim i As Long
    Dim ansi() As Byte
    Dim ascii As Integer
    Dim encText As String

    ansi = StrConv(urlText, vbFromUnicode)
    encText = ""
    For i = 0 To UBound(ansi)
        ascii = ansi(i)
        Select Case ascii
        Case 48 To 57, 65 To 90, 97 To 122
            encText = encText & Chr$(ascii)
        Case 32
            encText = encText & "+"
        Case Else
            If ascii < 16 Then
                encText = encText & "%0" & Hex$(ascii)
            Else
                encText = encText & "%" & Hex$(ascii)
            End If
        End Select
    Next i
    
    URLEncode = encText
End Function

Public Function Rot39(ByVal sData As String) As String

  'ROT39 (a variation of the ROT13 function) by Dag Sunde

   Dim sReturn As String
   Dim nCode As Long
   Dim nData As Long
   Dim bData() As Byte
   
   On Error GoTo Rot39_error
   
  'initialize the byte array to the
  'size of the string passed.
   ReDim bData(Len(sData)) As Byte
    
  'cast string into the byte array
   bData = StrConv(sData, vbFromUnicode)
    
   For nData = 0 To UBound(bData)
    
     'with the ASCII value of the character
      nCode = bData(nData)
        
     'assure the ASCII value is between
     'the lower and upper limits
      If ((nCode >= LOWER_LIMIT) And (nCode <= UPPER_LIMIT)) Then
         
        'shift the ASCII value by the
        'CHARMAP const value
         nCode = nCode + CHARMAP
         
        'perform a check against the upper
        'limit. If the new value exceeds the
        'upper limit, rotate the value to offset
        'from the beginning of the character set.
         If nCode > UPPER_LIMIT Then
            nCode = nCode - UPPER_LIMIT + LOWER_LIMIT - 1
         End If
      End If
        
     'reassign the new shifted value to
     'the current byte
      bData(nData) = nCode
        
   Next nData
    
  'convert the byte array back
  'to a string and exit
   sReturn = StrConv(bData, vbUnicode)

Rot39_exit:
   
  'assign the return string
   Rot39 = sReturn
   Exit Function
    
Rot39_error:
 
  'error! Return an empty string
   sReturn = ""
   Resume Rot39_exit:
    
End Function

Public Sub ShowBalloonTip(ByVal Title As String, ByVal Text As String, hWnd As Long, Optional icon As Long = 1)

   Dim ebt As EDITBALLOONTIP
   
   With ebt
      .cbStruct = Len(ebt)
      .pszTitle = StrConv(Title, vbUnicode)
      .pszText = StrConv(Text, vbUnicode)
      .ttiIcon = icon
   End With

   Call SendMessage(hWnd, EM_SHOWBALLOONTIP, 0&, ebt)
 
End Sub
Public Sub HideBallonTip(hWnd As Long)
    Call SendMessage(hWnd, EM_HIDEBALLOONTIP, 0&, 0&)
 
End Sub
Public Sub RegRun(Path As String, KeyName As String)
   ' Dim Reg As Object
   ' Set Reg = CreateObject("wscript.shell")
   ' Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & Keyname, Path
    Call SaveString(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", KeyName, Path)
    
    
End Sub
Public Sub RemoveRegRun(KeyName As String)
    Call DelSetting(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", KeyName)
End Sub



Public Sub AddToErrorLog(ByVal strText As String)
    Dim f As Long
    f = FreeFile
    Open App.Path & "\log.txt" For Append As #f
        Print #f, strText
    Close #f
End Sub
Function GetPost(Url As String, ByVal PostData As String) As String
On Error GoTo errHandle:
    Dim xmlhttp As Object
    
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Indicate that page that will receive the request and the
    ' type of request being submitted
    xmlhttp.Open "POST", Url, False
    
    ' Indicate that the body of the request contains form data
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the data as name/value pairs

    xmlhttp.send Trim$(PostData)

    GetPost = xmlhttp.responseText

    Set xmlhttp = Nothing
Exit Function
errHandle:
    MsgBox "Error: GetPost: - " & Err.Description
End Function
Function GetUrl(Url As String) As String
On Error GoTo errHandle:
    Dim xmlhttp As Object
    
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Indicate that page that will receive the request and the
    ' type of request being submitted
    xmlhttp.Open "GET", Url, False
    
    ' Indicate that the body of the request contains form data

    ' Send the data as name/value pairs

    xmlhttp.send

    GetUrl = xmlhttp.responseText

    Set xmlhttp = Nothing
Exit Function
errHandle:
    MsgBox "Error: GetUrl: - " & Err.Description
End Function
