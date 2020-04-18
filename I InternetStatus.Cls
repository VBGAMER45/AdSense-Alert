VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IInternetStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' IInternetStatus Interface
'
' This is implemented by classes that support asynchronous WinInet calls.
'
'===========================================================================
'
' Author:             Monte Hansen [monte@killervb.com]
' Dependencies:       MWinInet [M WinInet.Bas]
'
'===========================================================================
'
' == Copyright © 1999-2001 by Monte Hansen, All Rights Reserved Worldwide ==
'
' Monte Hansen (The Author) grants a royalty-free right to use,  modify, and
' distribute this code (The Code) in compiled form,  provided that you agree
' that The Author has no warranty,  obligations or liability  for  The Code.
' You may distribute The Code among peers but may not sell it, or distribute
' it on any electronic or physical media such  as  floppy diskettes, compact
' disks, bulletin boards, web sites,  and the like,  without first obtaining
' The Author's consent.
'
' When distributing The Code among peers,  it is respectfully requested that
' it be distributed as is,  but  at  no time shall it be distributed without
' the copyright notice hereinabove.
'
'===========================================================================
Option Explicit

' //
' // status manifests for Internet status callback
' //
Public Enum InternetStatus
  ' This is a custom notifier to send keep-alives:
  SendKeepAlive = -1
  ' These are standard status notifiers:
  ResolvingName = 10
  NameResolved = 11
  ConnectingToServer = 20
  ConnectedToServer = 21
  SendingRequest = 30
  RequestSent = 31
  ReceivingResponse = 40
  ResponseReceived = 41
  'ControlResponseReceived = 42
  PreFetch = 43
  ClosingConnection = 50
  ConnectionClosed = 51
  HandleCreated = 60
  HandleClosing = 70
  DetectingProxy = 80
  RequestComplete = 100
  Redirecting = 110
  IntermediateResponse = 120
  UserInputRequired = 140
  StateChange = 200
End Enum

Public Sub Callback( _
  ByVal dwInternetStatus As InternetStatus, _
  ByVal lpvStatusInformation As Long, _
  ByVal dwStatusInformationLength As Long)
  '
End Sub

