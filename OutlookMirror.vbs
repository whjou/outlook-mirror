' ------------------------------------
' OlInspectorClose http://msdn.microsoft.com/en-us/library/office/ff867882(v=office.14).aspx
Const olDiscard = 1 ' Changes to the document are discarded.
Const olSave    = 0 ' Documents are saved.



' ------------------------------------
' OlDefaultFolders http://msdn.microsoft.com/en-us/library/office/ff861868(v=office.14).aspx
Const olFolderInbox  = 6 ' The Inbox folder.
Const olFolderOutbox = 4 ' The Outbox folder.



' ------------------------------------
' OlObjectClass http://msdn.microsoft.com/en-us/library/bb208118(v=office.12).aspx
Const olFolder = 2  ' Represents a Folder object.
Const olMail   = 43 ' Represents a MailItem object.



' ------------------------------------
' Log Levels
Const llTRACE = 0
Const llDEBUG = 1
Const llINFO  = 2
Const llWARN  = 3
Const llERROR = 4
Dim llLabels(4)
llLabels(llTRACE) = "TRACE"
llLabels(llDEBUG) = "DEBUG"
llLabels(llINFO)  = "INFO "
llLabels(llWARN)  = "WARN "
llLabels(llERROR) = "ERROR"



' ------------------------------------
sResumeSourceFolderPath = ""
dtStartDateTime         = CDate("1900-01-01 00:00:00")
dtEndDateTime           = CDate("2100-01-01 00:00:00")
sMailItemFilter         = ""
iCount                  = 0



' ------------------------------------
Public oOutlook
Public oNameSpace
Public iLogLevel
iLogLevel = llDEBUG



' ------------------------------------
Sub olInit()
  Set oOutlook   = WScript.CreateObject("Outlook.Application")
  Set oNameSpace = oOutlook.GetNamespace("MAPI")
End Sub



' ------------------------------------
Function olGetFolder(ByVal sFolderPath)
  On Error Resume Next
  Dim aFolderPath
  Dim oFolder
  Dim i
  sFolderPath = Replace(sFolderPath, "/", "\")
  If Left(sFolderPath, 2) = "\\" Then
    sFolderPath = Right(sFolderPath, Len(sFolderPath) - 2)
  End If
  aFolderPath = Split(sFolderPath, "\")
  Set oFolder = oNameSpace.Folders(aFolderPath(0))
  For i = 1 To UBound(aFolderPath)
    Set oFolder = oFolder.Folders(aFolderPath(i))
    If Err <> 0 Then
      Set oFolder = Nothing
      Exit For
    End If
  Next
  Set olGetFolder = oFolder
  On Error Goto 0
End Function



' ------------------------------------
Function sQuote(s)
  sQuote = Chr(34) & Replace(Replace(s, "'", "''"), """", """""") & Chr(34)
End Function



' ------------------------------------
Function sFormatDate(dtDate)
  yyyy        = DatePart("yyyy", dtDate)
  m           = Right("00" & DatePart("m", dtDate), 2)
  d           = Right("00" & DatePart("d", dtDate), 2)
  sFormatDate = yyyy & "-" & m & "-" & d
End Function



' ------------------------------------
Function sFormatTime(dtDate)
  h           = Right("00" & DatePart("h", dtDate), 2)
  n           = Right("00" & DatePart("n", dtDate), 2)
  s           = Right("00" & DatePart("s", dtDate), 2)
  sFormatTime = h & ":" & n & ":" & s
End Function



' ------------------------------------
Function sFormatDateTime(dtDate)
  sFormatDateTime = sFormatDate(dtDate) & " " & sFormatTime(dtDate)
End Function



' ------------------------------------
Function sFormatFilterDate(dtDate)
  yy                = Right(DatePart("yyyy", dtDate),2)
  m                 = DatePart("m", dtDate)
  d                 = DatePart("d", dtDate)
  sFormatFilterDate = m & "/" & d & "/" & yy
End Function



' ------------------------------------
Function sFormatFilterTime(dtDate)
  h = DatePart("h", dtDate)
  n = Right("00" & DatePart("n", dtDate), 2)
  p = "AM"
  If h >= 12 Then
    p = "PM"
  End If
  sFormatFilterTime = h & ":" & n & " " & p
End Function



' ------------------------------------
Function sFormatFilterDateTime(dtDate)
  sFormatFilterDateTime = sFormatFilterDate(dtDate) & " " & sFormatFilterTime(dtDate)
End Function



' ------------------------------------
Sub LogLevel(ll)
  iLogLevel = ll
End Sub

' ------------------------------------
Sub Log(ll, s)
  If (ll >= iLogLevel) Then
    WScript.Echo "[" & sFormatDateTime(Now) & "] : " & llLabels(ll) & " : " & s
  End If
End Sub


' ------------------------------------
Sub LogTrace(s)
  Log llTRACE, s
End Sub



' ------------------------------------
Sub LogDebug(s)
  Log llDEBUG, s
End Sub



' ------------------------------------
Sub LogInfo(s)
  Log llINFO, s
End Sub



' ------------------------------------
Sub LogWarn(s)
  Log llWARN, s
End Sub



' ------------------------------------
Sub LogError(s)
  Log llERROR, s
End Sub



' ------------------------------------
Function olMailItemLong(oMailItem)
  On Error Resume Next
  s = "MailItem"
  s = s & " [Categories="           & oMailItem.Categories           & "]"
  s = s & " [CC="                   & oMailItem.CC                   & "]"
  s = s & " [ConversationTopic="    & oMailItem.ConversationTopic    & "]"
  s = s & " [Importance="           & oMailItem.Importance           & "]"
  s = s & " [LastModificationTime=" & oMailItem.LastModificationTime & "]"
  s = s & " [ReceivedTime="         & oMailItem.ReceivedTime         & "]"
  s = s & " [Recipients="           & oMailItem.Recipients           & "]"
  s = s & " [SenderEmailAddress="   & oMailItem.SenderEmailAddress   & "]"
  s = s & " [SenderName="           & oMailItem.SenderName           & "]"
  s = s & " [SentOn="               & oMailItem.SentOn               & "]"
  s = s & " [SentOnBehalfOfName="   & oMailItem.SentOnBehalfOfName   & "]"
  s = s & " [Size="                 & oMailItem.Size                 & "]"
  s = s & " [Subject="              & oMailItem.Subject              & "]"
  s = s & " [To="                   & oMailItem.To                   & "]"
  On Error Goto 0
  olMailItemLong = s
End Function



' ------------------------------------
Function olMailItemShort(oMailItem)
  On Error Resume Next
  s = "MailItem"
  s = s & " [SenderName="           & oMailItem.SenderName           & "]"
  s = s & " [SentOn="               & oMailItem.SentOn               & "]"
  s = s & " [Subject="              & oMailItem.Subject              & "]"
  On Error Goto 0
  olMailItemShort = s
End Function



' ------------------------------------
Function olMailItemExist(oFolder, oMailItem)
  Set c = oFolder.Items

  On Error Resume Next

  ' Filter by [SentOn]
  dtSentOn = oMailItem.SentOn
  sFilter  = "[SentOn]>=" & sQuote(sFormatFilterDateTime(dtSentOn)) & " AND [SentOn]<" & sQuote(sFormatFilterDateTime(dtSentOn + 30))
  Set c = c.Restrict(sFilter)
  LogTrace sFilter & " : " & c.Count

  ' Filter by [SenderName]
  sSenderName = oMailItem.SenderName
  If sSenderName <> "" Then
    sFilter     = "[SenderName]=" & sQuote(sSenderName)
    Set c = c.Restrict(sFilter)
    LogTrace sFilter & " : " & c.Count
  End If

  ' Filter by [Subject]
  sSubject = oMailItem.Subject
  If sSubject <> "" Then
    sSubjectPrefix = Left(oMailItem.Subject, 4)
    If sSubjectPrefix = "RE: " Or sSubjectPrefix = "FW: " Then
      sSubject = Mid(sSubject, 5)
    End If
    sFilter = "[Subject]=" & sQuote(sSubject)
    Set c = c.Restrict(sFilter)
    LogTrace sFilter & " : " & c.Count
  End If

  On Error Goto 0

  If c.Count = 0 Then
    olMailItemExist = False
  Else
    If iLogLevel <= llTRACE Then
      For Each oItem in c
        LogTrace "Exist : " & olMailItemShort(oItem)
      Next
    End If
    olMailItemExist = True
  End If
End Function



' ------------------------------------
Sub olMirror(sStartSourceFolderPath, oStartSourceFolder, oStartTargetFolder)
  LogInfo "BEGIN"
  olRecursiveCopyFolder sStartSourceFolderPath, oStartSourceFolder, oStartTargetFolder
  LogInfo "END : " & iCount & " items copied in total."
End Sub



' ------------------------------------
Sub olRecursiveCopyFolder(ByVal sCurrentFolderPath, oSourceFolder, oTargetFolder)
  LogInfo sCurrentFolderPath & " : BEGIN : " & oSourceFolder.Items.Count & " items; " & oSourceFolder.Folders.Count & " folders"
  olCopyFolder sCurrentFolderPath, oSourceFolder, oTargetFolder
  olCopySubFolders sCurrentFolderPath, oSourceFolder, oTargetFolder
  LogInfo sCurrentFolderPath & " : END"
End Sub



' ------------------------------------
Sub olCopyFolder(sCurrentFolderPath, oSourceFolder, oTargetFolder)
  'On Error Resume Next
  'lCount = oSourceFolder.Items.Count
  'oSourceFolder.CopyTo(oTargetFolder)
  'On Error Goto 0

  lCount = 0
  LogInfo sCurrentFolderPath & " : Contains " & oSourceFolder.Items.Count & " items."
  Set cSourceFolderFilteredItems = oSourceFolder.Items
  If sMailItemFilter <> "" Then
    Set cSourceFolderFilteredItems = cSourceFolderFilteredItems.Restrict(sMailItemFilter)
  End If
  oTargetFolder.Items.SetColumns("SenderName, Subject, ReceivedTime, Size")
  LogInfo sCurrentFolderPath & " : Copying  " & cSourceFolderFilteredItems.Count & " items."
  For Each oMailItem in cSourceFolderFilteredItems
    If olMailItemExist(oTargetFolder, oMailItem) Then
      ' Skip duplicate
      WScript.StdOut.Write "x"
      LogDebug "Skip : " & olMailItemShort(oMailItem)
      LogTrace "Skip : " & olMailItemLong(oMailItem)
    Else
      WScript.StdOut.Write "o"
      LogDebug "Copy : " & olMailItemShort(oMailItem)
      LogTrace "Copy : " & olMailItemLong(oMailItem)
      Set oCopy = oMailItem.Copy
      oMailItem.Close olDiscard
      oCopy.Move oTargetFolder
      oCopy.Close olDiscard
      Set oCopy = Nothing
      lCount = lCount + 1
    End If
  Next
  WScript.StdOut.WriteLine ""
  iCount = iCount + lCount
  LogInfo sCurrentFolderPath & " : " & lCount & " items copied; " & iCount & " items copied in total so far."
End Sub



' ------------------------------------
Sub olCopySubFolders(ByVal sCurrentFolderPath, oSourceFolder, oTargetFolder)
  For Each oSourceSubFolder in oSourceFolder.Folders
    On Error Resume Next
    Set oTargetSubFolder = oTargetFolder.Folders(oSourceSubFolder.Name)
    If Err <> 0 Then
      Set oTargetSubFolder = oTargetFolder.Folders.Add(oSourceSubFolder.Name)
    End If
    On Error Goto 0
    olRecursiveCopyFolder(sCurrentFolderPath & "\" & oSourceSubFolder.Name), oSourceSubFolder, oTargetSubFolder
  Next
End Sub



' ------------------------------------
Set oArgs = WScript.Arguments
If oArgs.Count < 4 Then
  WScript.Echo "Usage:"
  WScript.Echo "    " & WScript.ScriptName & " sourceFolder targetFolder fromDateTime toDateTime"
  WScript.Echo "Example:"
  WScript.Echo "    " & WScript.ScriptName & " ""Mailbox - User\Inbox"" ""Archive\Inbox"" - ""2014-12-15 00:00:00"""
  WScript.Quit(1)
End If

sStartSourceFolderPath = oArgs(0)
sStartTargetFolderPath = oArgs(1)
sStartDateTime         = oArgs(2)
sEndDateTime           = oArgs(3)

sMailItemFilter = ""
If IsDate(sStartDateTime) Then
  If Len(sStartDateTime) = 10 Then
    dtStartDateTime = CDate(sStartDateTime & " 00:00:00")
  Else
    dtStartDateTime = CDate(sStartDateTime)
  End If
  sMailItemFilter = "[SentOn]>=" & sQuote(sFormatFilterDateTime(dtStartDateTime))
End If
If IsDate(sEndDateTime) Then
  If Len(sStartDateTime) = 10 Then
    dtEndDateTime = CDate(sEndDateTime & " 00:00:00")
  Else
    dtEndDateTime = CDate(sEndDateTime)
  End If
  If sMailItemFilter <> "" Then
    sMailItemFilter = sMailItemFilter & " AND "
  End If
  sMailItemFilter = sMailItemFilter & "[SentOn]<" & sQuote(sFormatFilterDateTime(dtEndDateTime))
End If

olInit

LogInfo "Source : " & sStartSourceFolderPath
LogInfo "Target : " & sStartTargetFolderPath
If sMailItemFilter <> "" Then
  LogInfo "Filter : " & sMailItemFilter
End If

Set oStartSourceFolder = olGetFolder(sStartSourceFolderPath)
Set oStartTargetFolder = olGetFolder(sStartTargetFolderPath)
If oStartSourceFolder Is Nothing Or oStartTargetFolder Is Nothing Then
  LogError "Missing start source or target folder"
  WScript.Quit(1)
End If

olMirror sStartSourceFolderPath, oStartSourceFolder, oStartTargetFolder

WScript.Quit(1)
