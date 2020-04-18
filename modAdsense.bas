Attribute VB_Name = "modAdsense"
Option Explicit
#Const Trial = False

Public Sub GetTodaysAdData(lst As ListView, Optional dateRange As String = "today", Optional UpdateList As Boolean = True, Optional IsSearch As Boolean = False)
On Error GoTo errHandle
    If TrialExpired = True Then
        Exit Sub
    End If

        Call modGlobals.CheckTrial


       Dim strData As String


       Dim Final As String

        If IsSearch = False Then
                'If strChannel = "" Then
               ' strData = GetPost("https://www.google.com/adsense/login.do", "destination=%2Fadsense%2Freports-aggregate%3Fproduct%3Dafc%26dateRange.dateRangeType%3Dsimple%26dateRange.simpleDate%3D" & dateRange & "%26reportType%3Dproperty%26csv%3Dtrue&username=" & UserName & "&password=" & Password)
                strData = GetPost("https://www.google.com/adsense/report/aggregate?product=afc&dateRange.dateRangeType=simple&dateRange.simpleDate=today&dateRange.customDate.start.month=" & frmMain.cboMonth1.Text & "&dateRange.customDate.start.day=1&dateRange.customDate.start.year=" & frmMain.cboYear1.Text & "&dateRange.customDate.end.month=" & frmMain.cboMonth2.Text & "&dateRange.customDate.end.day=" & frmMain.cboDay2.Text & "&dateRange.customDate.end.year=" & frmMain.cboYear2.Text & "&groupByPref=date&unitPref=page&reportType=property&null=Display+Report&csv=true", "username=" & UsernameSafe & "&password=" & PasswordSafe)
               Debug.Print strData
                'Else
                'strData = h.SendRequest("/adsense/login.do?destination=%2Fadsense%2Freports-aggregate%3Fproduct%3Dafc%26dateRange.dateRangeType%3Dsimple%26dateRange.simpleDate%3D" & dateRange & "%26reportType%3Dchannel" & strChannel & "%26csv%3Dtrue&username=" & UserName & "&password=" & Password, "GET")
                'End If
        Else
            strData = GetPost("https://www.google.com/adsense/report/aggregate?product=afs", "dateRange.dateRangeType%3Dsimple%26dateRange.simpleDate%3D" & dateRange & "%26reportType%3Dproperty%26csv%3Dtrue&username=" & UsernameSafe & "&password=" & PasswordSafe)
            Debug.Print strData
          '  strData = GetPost("https://www.google.com/adsense/login.do", "destination=%2Fadsense%2Freports-aggregate%3Fproduct%3Dafs%26dateRange.dateRangeType%3Dsimple%26dateRange.simpleDate%3D" & dateRange & "%26reportType%3Dproperty%26csv%3Dtrue&username=" & UserName & "&password=" & Password)
        End If
          
            Final = StrConv(Mid$(strData, 3), vbFromUnicode)
            Dim Temp() As String, Temp2() As String
            Temp = Split(Final, vbLf)
            Dim i As Long, Count As Long
            If Final = "" Then Exit Sub
            
            If dateRange = "today" Then
                Temp2 = Split(Temp(1), vbTab)
                TodayClicks = Temp2(2)
                TodayImpressions = Temp2(1)
                TodayCTR = CSng(Replace(Temp2(3), "%", ""))
                TodayCPM = CSng(Temp2(2))
                TodayEarnings = Temp2(5)
               '###JV Call frmMain.AddPopup("AdseneAlert", "Today's Earnings: " & TodayEarnings & vbCrLf & "Today's Clicks: " & TodayClicks & vbCrLf & "Today Impressions: " & TodayImpressions & vbCrLf & "Today's CTR: " & Temp2(3) & vbCrLf & "Today's CPM: " & Temp2(4))
                If EmailUpdates = True Then
                    Call SendReportEmail("Adsense Alert Update for " & Temp2(0) & vbCrLf & "Today's Earnings: " & TodayEarnings & vbCrLf & "Today's Clicks: " & TodayClicks & vbCrLf & "Today Impressions: " & TodayImpressions & vbCrLf & "Today's CTR: " & Temp2(3) & vbCrLf & "Today's CPM: " & Temp2(4))
                    
                End If
                Call frmMain.UpdateStatusBar
                frmMain.tmrAlertLoop.Enabled = True
            End If
            If UpdateList = True Then
                lst.ListItems.Clear
                For i = 1 To UBound(Temp) - 1
    
          
                      Temp2 = Split(Temp(i), vbTab)
                      lst.ListItems.Add , , Temp2(0)
                      Count = Count + 1
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(1)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(2)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(3)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(4)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(5)
                  
                Next
                'Do averages
                lst.ListItems.Item(Count).ForeColor = vbBlue
                lst.ListItems.Add , , "Averages"
                Count = Count + 1
                lst.ListItems.Item(Count).ForeColor = vbRed
                lst.ListItems.Item(Count).ListSubItems.Add , , CLng(Temp2(1) / (Count - 2))
                lst.ListItems.Item(Count).ListSubItems.Add , , CLng(Temp2(2) / (Count - 2))
                'lst.ListItems.Item(count).ListSubItems.Add , , Format(CSng(Replace(Temp2(3), "%", "")) / CSng(count - 2), "###.##")
                'lst.ListItems.Item(count).ListSubItems.Add , , Format(CSng(Temp2(4)) / CSng((count - 2)), "###.##")
                lst.ListItems.Item(Count).ListSubItems.Add , , ""
                lst.ListItems.Item(Count).ListSubItems.Add , , ""
                lst.ListItems.Item(Count).ListSubItems.Add , , Format(CSng(Temp2(5)) / CSng((Count - 2)), "###.##")
                  
            End If

        
'https://www.google.com/adsense/report/aggregate?product=afc&dateRange.dateRangeType=simple&dateRange.simpleDate=yesterday&dateRange.customDate.start.month=6&dateRange.customDate.start.day=21&dateRange.customDate.start.year=2005&dateRange.customDate.end.month=6&dateRange.customDate.end.day=21&dateRange.customDate.end.year=2005&groupByPref=date&unitPref=page&reportType=property&null=Display+Report
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_GetTodaysAdData : " & Err.Description & " " & Date & " " & Time)
End Sub


Public Sub GetDateRange(lst As ListView, Month1 As Integer, day1 As Integer, year1 As Integer, month2 As Integer, day2 As Integer, year2 As Integer)
On Error GoTo errHandle
    If TrialExpired = True Then
        Exit Sub
    End If

    Call modGlobals.CheckTrial


    Dim strData As String

       Dim Final As String
          strData = GetPost("https://www.google.com/adsense/login.do", "destination=%2Fadsense%2Freports-aggregate%3Fproduct%3Dafc%26dateRange.dateRangeType%3Dsimple%26dateRange.simpleDate%3Dtoday%26dateRange.dateRangeType%3Dcustom%26dateRange.customDate.start.month%3D" & Month1 & "%26dateRange.customDate.start.day%3D" & day1 & "%26dateRange.customDate.start.year%3D" & year1 & "%26dateRange.customDate.end.month%3D" & month2 & "%26dateRange.customDate.end.day%3D" & day2 & "%26dateRange.customDate.end.year%3D" & year2 & "%26groupByPref%3Ddate%26unitPref%3Dpage%26reportType%3Dproperty%26null%3DDisplay%2BReport%26csv%3Dtrue&username=" & UserName & "&password=" & Password)
            Final = StrConv(Mid$(strData, 3), vbFromUnicode)
            Dim Temp() As String, Temp2() As String
            Temp = Split(Final, vbLf)
            Dim i As Long, Count As Long
            
            

                lst.ListItems.Clear
                For i = 1 To UBound(Temp) - 1
    
          
                    Temp2 = Split(Temp(i), vbTab)
                    lst.ListItems.Add , , Temp2(0)
                    Count = Count + 1
                    lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(1)
                    lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(2)
                    lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(3)
                    lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(4)
                    lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(5)
                  
                Next


    
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_GetDateRange : " & Err.Description & " " & Date & " " & Time)

End Sub
Public Sub GetPayment(lst As ListView, Month1 As Integer, month2 As Integer, year1 As Integer, year2 As Integer)
On Error GoTo errHandle



       Dim strData As String



       Dim Final As String
        'strData = h.SendRequest("/adsense/login.do", "GET")
           
           strData = GetPost("https://www.google.com/adsense/login.do", "destination=%2Fadsense%2Freports-payment%3Fbegin.month%3D" & Month1 & "%26begin.year%3D" & year1 & "%26end.month%3D" & month2 & "%26end.year%3D" & year2 & "%26null%3DGo%26csv%3Dtrue&username=" & UserName & "&password=" & Password)
           'stdata = h.SendRequest("%2Fadsense%2Freports-payment%3Fbegin.month%3D1%26begin.year%3D2004%26end.month%3D2%26end.year%3D2005%26null%3DGo%26csv%3Dtrue&username=" & UserName & "&password=" & Password, "GET")

           
            'strData = h.SendRequest("/adsense/login.do?destination=%2Fadsense%2Freports-payment%3Fbegin.month%3D" & Month1 & "%26begin.year%3D" & year1 & "%26end.month%3D" & month2 & "%26end.year%3D" & year2 & "%26null%3DGo%26csv%3Dtrue&username=" & UserName & "&password=" & Password, "GET")
            ''h.Fields("username") = UserName
            ''h.Fields("password") = Password
            ''h.Fields("destination") = "/adsense/reports-payment?begin.month=" & Month1 & "&begin.year=" & year1 & "&end.month=" & month2 & "&end.year=" & year2 & "&null=Go&csv=true"
            
            ''strData = h.SendRequest("/adsense/login.do") '?destination=/adsense/reports-payment?begin.month=" & Month1 & "&begin.year=" & year1 & "&end.month=" & month2 & "&end.year=" & year2 & "&null=Go&csv=true", "GET")

            Final = StrConv(Mid$(strData, 3), vbFromUnicode)
          '  MsgBox Final
            Dim Temp() As String, Temp2() As String
            Temp = Split(Final, vbLf)
            Dim i As Long, Count As Long

                lst.ListItems.Clear
            For i = 1 To UBound(Temp) - 1
    
          
                Temp2 = Split(Temp(i), vbTab)
                lst.ListItems.Add , , Temp2(0)
                Count = Count + 1
                lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(1)
                lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(2)

                
            Next




        

        
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_GetPayment : " & Err.Description & " " & Date & " " & Time)
End Sub
Public Sub GetUrlChannels(strData As String)
On Error GoTo errHandle
    Dim strSearch As String

    strSearch = "<optgroup label=" & Chr$(34) & "Active URL Channels:" & Chr$(34) & ">"
    Dim pos As Long, pos2 As Long
    pos = InStr(1, strData, strSearch)
    pos2 = InStr(1, strData, "</optgroup>")
    ReDim ChannelList(0)
    If pos <> 0 And pos2 <> 0 Then
        Dim Data As String
        Data = Mid$(strData, pos + Len(strSearch), pos2 - pos - Len(strSearch))
        If InStr(1, Data, "</option>") <> 0 Then
            Dim Temp() As String
            Temp = Split(Data, "</option>")
            Dim i As Long
            For i = 0 To UBound(Temp)
                Dim pos3 As Long, pos4 As Long
                pos3 = InStr(1, Temp(i), "value=")
                pos4 = InStr(1, Temp(i), " class=")
                If pos3 <> 0 And pos4 <> 0 Then
                    Dim k As String
                    'Get The value
                    ChannelList(UBound(ChannelList)).ChannelId = Mid$(Temp(i), pos3 + 7, pos4 - pos3 - 8)
                    pos3 = InStr(1, Temp(i), "title=")
                    pos4 = InStr(1, Temp(i), "> ")
                    ChannelList(UBound(ChannelList)).ChannelName = Mid$(Temp(i), pos3 + 7, pos4 - pos3 - 8)
                    ReDim Preserve ChannelList(UBound(ChannelList) + 1)
                End If
            Next
            
        End If
        ReDim Preserve ChannelList(UBound(ChannelList) - 1)
        
       ' Debug.Print Mid$(strData, pos + Len(strSearch), pos2 - pos - Len(strSearch))
    End If
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_GetUrlChannels : " & Err.Description & " " & Date & " " & Time)

End Sub

Public Sub PrintAdReport(lst As ListView, dateRange As String)
On Error GoTo errHandle
    Printer.FontBold = True
    Printer.FontSize = 24
    Printer.Print "AdsenseAlert.com - Report  " & lst.ListItems.Item(1).Text & " - " & lst.ListItems.Item(lst.ListItems.Count - 2).Text
    Printer.Print vbCrLf
    Printer.FontSize = 12
    Printer.Print "Date" & vbTab & "Page Impressions" & vbTab & "Clicks" & vbTab & "Page CTR" & vbTab & "Page eCPM" & vbTab & "Your earnings"
    Printer.FontBold = False
    Dim i As Long
    If lst.ListItems.Count > 1 Then
        For i = 1 To lst.ListItems.Count - 1
            Printer.Print lst.ListItems.Item(i).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(1).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(2).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(3).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(4).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(5).Text
        Next
    End If
    i = lst.ListItems.Count
    Printer.Print lst.ListItems.Item(i).Text & vbTab & lst.ListItems.Item(i).ListSubItems(1).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(2).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(3).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(4).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(5).Text
    
    Printer.EndDoc
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_PrindAdReport : " & Err.Description & " " & Date & " " & Time)

End Sub
Public Sub PrintPayments(lst As ListView)
On Error GoTo errHandle
    Printer.FontBold = True
    Printer.FontSize = 24
    Printer.Print "AdsenseAlert.com - Report  " & lst.ListItems.Item(1).Text & " - " & lst.ListItems.Item(lst.ListItems.Count).Text
    Printer.FontSize = 20
    Printer.Print "Payment History"
    Printer.Print vbCrLf
    Printer.FontSize = 12
    Printer.Print "Date" & vbTab & "Description" & vbTab & vbTab & "Amount"
    Printer.FontBold = False
    Dim i As Long
    'If lst.ListItems.count > 1 Then
        For i = 1 To lst.ListItems.Count
            Printer.Print lst.ListItems.Item(i).Text & vbTab & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(1).Text & vbTab & vbTab & lst.ListItems.Item(i).ListSubItems(2).Text
        Next
   ' End If
 
    Printer.EndDoc
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_PrintPayments : " & Err.Description & " " & Date & " " & Time)

End Sub

Public Sub GetTodaysChannelData(lst As ListView, Optional dateRange As String = "today", Optional strChannel As String = "")
On Error GoTo errHandle
    If TrialExpired = True Then
        Exit Sub
    End If


       Dim strData As String



       Dim Final As String

        'strData = h.SendRequest("/adsense/login.do?destination=%2Fadsense%2Freports-aggregate%3Fproduct%3Dafc%26dateRange.dateRangeType%3Dsimple%26dateRange.simpleDate%3D" & dateRange & "&groupByPref=date%26reportType%3Dchannel" & strChannel & "%26csv%3Dtrue&null=Display+Report&username=" & UserName & "&password=" & Password, "GET")
       strData = GetPost("https://www.google.com/adsense/login.do", "destination=%2Fadsense%2Freports-aggregate%3Fproduct=afc&dateRange.dateRangeType=simple&dateRange.simpleDate=today&dateRange.customDate.start.month=7&dateRange.customDate.start.day=1&dateRange.customDate.start.year=2005&dateRange.customDate.end.month=7&dateRange.customDate.end.day=1&dateRange.customDate.end.year=2005&groupByPref=date&reportType=channel&c.id=1096346&null=Display+Report&csv=true&username=" & UserName & "&password=" & Password)
       

 
          
            Final = StrConv(Mid$(strData, 3), vbFromUnicode)
            Dim Temp() As String, Temp2() As String
            Temp = Split(Final, vbLf)
            Dim i As Long, Count As Long
            
            

                lst.ListItems.Clear
                For i = 1 To UBound(Temp) - 1
    
          
                      Temp2 = Split(Temp(i), vbTab)
                      lst.ListItems.Add , , Temp2(0)
                      Count = Count + 1
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(1)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(2)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(3)
                      lst.ListItems.Item(Count).ListSubItems.Add , , Temp2(4)

                  
                Next



        

        
Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_GetChannelData : " & Err.Description & " " & Date & " " & Time)
End Sub
Public Sub SendReportEmail(Message As String)
On Error GoTo errHandle
    If TrialExpired = True Then
        Exit Sub
    End If


MsgBox "FIX THIS"

       Dim strData As String

       Dim Final As String
 
       Dim strEmail As String
       strEmail = EmailUrl
       Dim Temp() As String
       Temp = Split(strEmail, "/")
  
        strData = GetPost(strEmail, "mailbody=" & Message)
        If InStr(1, strData, "Failed to send email") <> 0 Then
            Call AddToErrorLog("Failed to send email!")
        End If
        


Exit Sub
errHandle:
    Call AddToErrorLog("Error_modAdsense_SendReportEmail : " & Err.Description & " " & Date & " " & Time)
End Sub
