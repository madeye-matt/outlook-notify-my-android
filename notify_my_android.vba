Public Function URLEncode( _
   StringToEncode As String, _
   Optional UsePlusRatherThanHexForSpace As Boolean = False _
) As String

  Dim TempAns As String
  Dim CurChr As Integer
  CurChr = 1

  Do Until CurChr - 1 = Len(StringToEncode)
    Select Case Asc(Mid(StringToEncode, CurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
        TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
      Case 32
        If UsePlusRatherThanHexForSpace = True Then
          TempAns = TempAns & "+"
        Else
          TempAns = TempAns & "%" & Hex(32)
        End If
      Case Else
        TempAns = TempAns & "%" & _
          Right("0" & Hex(Asc(Mid(StringToEncode, _
          CurChr, 1))), 2)
    End Select

    CurChr = CurChr + 1
  Loop

  URLEncode = TempAns
End Function

Public Sub NotifyMyAndroid(sEvent As String, sDescription As String, sPriority As String)
    Dim description As String
    Dim apikey As String
    Dim appName As String
    Dim url As String
    
    apikey = "<INSERT YOU NOTIFY MY ANDROID APIKEY HERE>"
    appName = URLEncode("<INSERT YOUR APP NAME HERE>")

    url = "http://www.notifymyandroid.com/publicapi/notify?apikey=" & apikey & "&application=" & appName & "&event=" & sEvent & "&description=" & sDescription & "&priority=" & sPriority
    
    ' MsgBox "url: " & url
    
    Dim objReq As WinHttp.WinHttpRequest
    Set objReq = New WinHttp.WinHttpRequest
    objReq.Option(WinHttpRequestOption_EnableRedirects) = True
    objReq.Open "GET", url, False
    objReq.Send

End Sub

Sub WatchNotificationMessageRule(Item As Outlook.MailItem)
    Dim strID As String
    Dim objMail As Outlook.MailItem
    Dim subject As String
    Dim sender As String
    Dim body As String
    Dim description As String
    
    strID = Item.EntryID
    
    Set objMail = Application.Session.GetItemFromID(strID)
    subject = objMail.subject
    sender = objMail.sender
    body = Left$(objMail.body, 500)
    
    description = URLEncode(subject) & "%0d" & URLEncode(body)
    
    Call NotifyMyAndroid(sender, description, "0")
    
    Set objMail = Nothing
End Sub


Private Sub Application_Reminder(ByVal Item As Object)
  Dim sEvent As String
  Dim sDescription As String
  ' create new outgoing message
   ' your reminder notification address
  sEvent = "Reminder: " & Item.subject
  ' must handle all 4 types of items that can generate reminders
  Select Case Item.Class
     Case olAppointment '26
        sDescription = _
          "Start: " & Item.Start & vbCrLf & _
          "End: " & Item.End & vbCrLf & _
          "Location: " & Item.Location & vbCrLf & _
          "Details: " & vbCrLf & Item.body
     Case olContact '40
        sDescription = _
          "Contact: " & Item.FullName & vbCrLfrLf & _
          "Phone: " & Item.BusinessTelephoneNumber & vbCrLf & _
          "Contact Details: " & vbCrLf & Item.body
      Case olMail '43
        sDescription = _
          "Due: " & Item.FlagDueBy & vbCrLf & _
          "Details: " & vbCrLf & Item.body
      Case olTask '48
        sDescription = _
          "Start: " & Item.StartDate & vbCrLf & _
          "End: " & Item.DueDate & vbCrLf & _
          "Details: " & vbCrLf & Item.body
  End Select
  sEvent = Module1.URLEncode(sEvent)
  sDescription = Module1.URLEncode(sDescription)
  
  ' send the message
  Call Module1.NotifyMyAndroid(sEvent, sDescription, "1")
End Sub




