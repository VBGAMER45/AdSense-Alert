VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTPClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ePort
   INTERNET_DEFAULT_HTTP_PORT = 80
  INTERNET_DEFAULT_HTTPS_PORT = 443
End Enum

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_SERVICE_HTTP = 3

Private Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_SECURE = &H800000
Private Const INTERNET_FLAG_FROM_CACHE = &H1000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_ASYNC = &H10000000

Private Const BUFFER_LENGTH As Long = 1024

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal Agent As String, ByVal AccessType As Long, ByVal ProxyName As String, ByVal ProxyBypass As String, ByVal Flags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal ServerName As String, ByVal ServerPort As Integer, ByVal UserName As String, ByVal Password As String, ByVal Service As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Boolean

Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hConnect As Long, ByVal Buffer As String, ByVal NumberOfBytesToRead As Long, NumberOfBytesRead As Long) As Boolean

Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal Verb As String, ByVal ObjectName As String, ByVal Version As String, ByVal Referer As String, ByVal AcceptTypes As Long, ByVal Flags As Long, Context As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal Headers As String, ByVal HeadersLength As Long, ByVal sOptional As String, ByVal OptionalLength As Long) As Boolean

Private hHTTP As Long
Private hConnection As Long

Private Const FIELDS_BUFFER_LENGTH As Long = 10
Private Const FIELDS_NAME_INDEX As Long = 0
Private Const FIELDS_VALUE_INDEX As Long = 1

Private DontEncode(255) As Boolean

Private FieldCount As Long
Private mFields() As String

'InternetCallback
Private Declare Sub CopyMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nBytes As Long)
Private Declare Sub InternetSetStatusCallback Lib "wininet.dll" ( _
     ByRef hInternet As Long, _
     ByRef lpfnInternetCallback As Long)

' // status manifests for Internet status callback
Private Enum InternetStatusFlags
  INTERNET_STATUS_RESOLVING_NAME = 10
  INTERNET_STATUS_NAME_RESOLVED = 11
  INTERNET_STATUS_CONNECTING_TO_SERVER = 20
  INTERNET_STATUS_CONNECTED_TO_SERVER = 21
  INTERNET_STATUS_SENDING_REQUEST = 30
  INTERNET_STATUS_REQUEST_SENT = 31
  INTERNET_STATUS_RECEIVING_RESPONSE = 40
  INTERNET_STATUS_RESPONSE_RECEIVED = 41
  INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42
  INTERNET_STATUS_PREFETCH = 43
  INTERNET_STATUS_CLOSING_CONNECTION = 50
  INTERNET_STATUS_CONNECTION_CLOSED = 51
  INTERNET_STATUS_HANDLE_CREATED = 60
  INTERNET_STATUS_HANDLE_CLOSING = 70
  INTERNET_STATUS_DETECTING_PROXY = 80
  INTERNET_STATUS_REQUEST_COMPLETE = 100
  INTERNET_STATUS_REDIRECT = 110
  INTERNET_STATUS_INTERMEDIATE_RESPONSE = 120
  INTERNET_STATUS_USER_INPUT_REQUIRED = 140
  INTERNET_STATUS_STATE_CHANGE = 200
End Enum

Public Property Let Fields(Name As String, Value As String)

   mFields(FIELDS_VALUE_INDEX, GetFieldIndex(Name, True)) = Value

End Property

Public Property Get Fields(Name As String) As String

   Dim l As Long
   
   l = GetFieldIndex(Name, False)
   If l > -1 Then
      Fields = mFields(FIELDS_VALUE_INDEX, l)
   End If

End Property

Public Function OpenHTTP(Server As String, Optional Port As ePort = INTERNET_DEFAULT_HTTP_PORT, Optional UserName As String, Optional Password As String) As Boolean
   
   CloseHTTP
 '  Dim result As Long
   hHTTP = InternetOpen("AdsenseAlert", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0&) ', INTERNET_FLAG_ASYNC)
   If hHTTP <> 0 Then
     'result = InternetSetStatusCallback(hHTTP, AddressOf INTERNET_STATUS_CALLBACK)
     
      hConnection = InternetConnect(hHTTP, Server, INTERNET_DEFAULT_HTTP_PORT, UserName, Password, INTERNET_SERVICE_HTTP, 0, 0)
      If hConnection <> 0 Then
         OpenHTTP = True
      Else
         InternetCloseHandle hHTTP
         hHTTP = 0
      End If
   End If
   
End Function

Public Sub CloseHTTP()
    
    If hConnection <> 0 Then
      InternetCloseHandle hConnection
    End If
    
    hConnection = 0
    
    If hHTTP Then
      InternetCloseHandle hHTTP
    End If
    
    hHTTP = 0

End Sub

Public Function SendRequest(ByVal File As String, Optional Method As String = "GET", Optional Referer As String, Optional Reload As Boolean = True) As String

   Dim hRequest As Long
   Dim r As Boolean
   Dim Buffer As String
   Dim Header As String
   Dim Request As String
   Dim POSTData As String
   Dim Response As String
   Dim Read As Long
   Dim Flags As Long
   
   Method = UCase$(Method)
   Request = BuildRequest
   Buffer = Space$(BUFFER_LENGTH)
   
   
   If Len(Request) > 0 Then
      If Method = "POST" Then
         Header = "Content-Type: application/x-www-form-urlencoded"
         POSTData = Request
      Else
         File = File & "?" & Request
      End If
   End If
   
   If Reload Then
      Flags = Flags Or INTERNET_FLAG_PRAGMA_NOCACHE Or INTERNET_FLAG_RELOAD
   End If
   
   hRequest = HttpOpenRequest(hConnection, Method, File, "HTTP/1.1", "", 0, Flags, 0)
   If hRequest <> 0 Then
      If HttpSendRequest(hRequest, Header, Len(Header), POSTData, Len(POSTData)) Then
         r = InternetReadFile(hRequest, Buffer, BUFFER_LENGTH, Read)
         While r And (Read <> 0)
            Response = Response & Left$(Buffer, Read)
            r = InternetReadFile(hRequest, Buffer, BUFFER_LENGTH, Read)
            DoEvents
         Wend
      End If
      InternetCloseHandle hRequest
   End If

   SendRequest = Response
   
End Function

Private Function GetFieldIndex(Name As String, Optional Add As Boolean) As Long

   Dim l As Long
   
   For l = 0 To FieldCount - 1
      If StrComp(Name, mFields(FIELDS_NAME_INDEX, l), vbTextCompare) = 0 Then
         GetFieldIndex = l
         Exit Function
      End If
   Next
   
   If Add Then
      If FieldCount = UBound(mFields, 2) Then
         ReDim Preserve mFields(1, UBound(mFields, 2) + FIELDS_BUFFER_LENGTH)
      End If
      mFields(FIELDS_NAME_INDEX, FieldCount) = Name
      GetFieldIndex = FieldCount
      FieldCount = FieldCount + 1
   Else
      GetFieldIndex = -1
   End If
   
End Function

Private Function BuildRequest() As String

   Dim l As Long
   Dim s As String
   
   For l = 0 To FieldCount - 1
      s = s & URLEncode(mFields(FIELDS_NAME_INDEX, l)) & "=" & URLEncode(mFields(FIELDS_VALUE_INDEX, l)) & "&"
   Next

   If Len(s) > 0 Then
      BuildRequest = Left$(s, Len(s) - 1)
   End If

End Function

Public Function URLEncode(Data As String) As String

   Dim l As Long
   Dim b() As Byte
   Dim s As String
   Dim c As String
   
   b = Data
   'This is fine for encoding small strings
   'To encode large ones I suggest you replace s with the String Class
   For l = 0 To UBound(b) Step 2
      If DontEncode(b(l)) Then
         s = s & Chr(b(l))
      Else
         c = Hex(b(l))
         While Len(c) < 2
            c = "0" & c
         Wend
         s = s & "%" & c
      End If
   Next
   
   URLEncode = s
   
End Function

Private Sub Class_Initialize()

   Dim l As Long

   ReDim mFields(1, FIELDS_BUFFER_LENGTH)

   For l = Asc("0") To Asc("9")
      DontEncode(l) = True
   Next
   For l = Asc("a") To Asc("z")
      DontEncode(l) = True
   Next
   For l = Asc("A") To Asc("Z")
      DontEncode(l) = True
   Next

End Sub

Private Sub Class_Terminate()

   Erase mFields
   
End Sub
'Public Sub INTERNET_STATUS_CALLBACK( _
'  ByVal hInternet As Long, _
'  ByVal dwContext As Long, _
'  ByVal dwInternetStatus As InternetStatusFlags, _
'  ByVal lpvStatusInformation As Long, _
'  ByVal dwStatusInformationLength As Long)
'===========================================================================
' INTERNET_STATUS_CALLBACK - Asynchronous ballback function for WinInet.
' This function assumes the caller supplies a pointer to its IInternetStatus
' interface. The call is then routed to the calling instance.
'===========================================================================
    
'  Dim InternetObj As IInternetStatus


'  On Error GoTo ErrHandler

  ' Sanity checks
'  If InvalidHandleValue(dwContext) Then GoTo ExitLabel
  
  ' Get owner of this status InternetObj.
  ' It can be any class that Implements
  ' the IInternetStatus interface.
  'Set InternetObj = PtrToObj(dwContext)

 ' Call InternetObj.Callback( _
   ' dwInternetStatus, _
  '  lpvStatusInformation, _
 '   dwStatusInformationLength)
  
'ExitLabel:

  'Exit Sub
 ' Resume
'ErrHandler:
  'Debug.Assert 0
 ' Resume ExitLabel

'End Sub
Private Function InvalidHandleValue(ByVal Handle As Long) As Boolean

    Const INVALID_HANDLE_VALUE = -1
    
    Select Case Handle
    Case 0, INVALID_HANDLE_VALUE
        InvalidHandleValue = True
    End Select
    
End Function
Private Function IsValidHandle(ByVal Handle As Long) As Boolean

    Const INVALID_HANDLE_VALUE = -1
    
    Select Case Handle
    Case 0, INVALID_HANDLE_VALUE
    Case Else
        IsValidHandle = True
    End Select
    
End Function
Private Function PtrToObj(ByVal lpObject As Long) As Object

    Dim WeakRef     As Object

    On Error GoTo ErrHandler

    ' Get the object pointer attached to this window's user data
    If Not IsValidHandle(lpObject) Then Exit Function

    ' Create the weak reference
    Call CopyMemoryAny(WeakRef, lpObject, 4)

    ' Return the object reference
    If Not WeakRef Is Nothing Then
        Set PtrToObj = WeakRef
    End If

    ' Kill the illegal reference
    'ZeroMemory ByVal ObjPtr(WeakRef), 4
    Call CopyMemoryAny(WeakRef, 0&, 4&)
    
    Exit Function
    Resume
ErrHandler:
    Set PtrToObj = Nothing
    Resume Next
    
End Function

