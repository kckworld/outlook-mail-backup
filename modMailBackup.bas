Option Explicit

Public Const BACKUP_BASE_PATH As String = "D:\backup\mail\"

' 메일에 저장할 UserProperty 이름 (이전 백업 파일 경로 기억용)
Private Const PROP_BACKUP_PATH As String = "BackupFilePath"

'===========================================
' MSG 저장
' 저장 구조:
'   Inbox: D:\backup\mail\YYYY\<Inbox 하위폴더들>\...
'   Sent : D:\backup\mail\YYYY\Sent\<...>\...
'
' 중복 제거 정책:
' - Inbox 루트에서 1회 백업된 메일이,
'   나중에 Inbox 하위 폴더로 이동되어 백업될 경우
'   "새 백업 저장이 성공한 뒤"에만 기존(루트) 백업 파일을 삭제
'
' BackupFilePath(UserProperty):
' - 마지막으로 성공 저장된 백업 파일 경로를 기록
'===========================================
Public Sub SaveMailAsMSG(mail As Outlook.MailItem, folderType As String, Optional relativeFolderPath As String = "")
    On Error GoTo ErrorHandler

    Dim mailTime As Date

    ' 1) 시간 결정
    If folderType = "Sent" Then
        If IsDate(mail.SentOn) And mail.SentOn > #1/1/1900# Then
            mailTime = mail.SentOn
        ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
            mailTime = mail.CreationTime
        Else
            mailTime = Now
        End If
    Else
        If IsDate(mail.ReceivedTime) And mail.ReceivedTime > #1/1/1900# Then
            mailTime = mail.ReceivedTime
        ElseIf IsDate(mail.CreationTime) And mail.CreationTime > #1/1/1900# Then
            mailTime = mail.CreationTime
        Else
            mailTime = Now
        End If
    End If

    ' 2) 폴더 미러링 경로
    Dim mirrorPath As String
    mirrorPath = ""

    If Trim(relativeFolderPath) <> "" Then
        mirrorPath = CleanFolderPath(relativeFolderPath) & "\"
    End If

    If folderType = "Sent" Then
        mirrorPath = "Sent\" & mirrorPath
    End If

    ' 3) 저장 경로: D:\backup\mail\YYYY\<mirrorPath>\
    Dim savePath As String
    savePath = BACKUP_BASE_PATH & Format(mailTime, "yyyy") & "\" & mirrorPath
    CreateFolderPath savePath

    ' 4) 파일명 생성 (260 제한 고려)
    Dim dateTimePart As String
    Dim senderPart As String
    Dim subjectPart As String
    Dim maxPathLength As Long
    Dim availableLength As Long

    maxPathLength = 260
    dateTimePart = Format(mailTime, "yyyymmdd_hhnnss")

    availableLength = maxPathLength - Len(savePath) - Len(dateTimePart) - 2 - 4 - 5
    If availableLength < 30 Then availableLength = 30

    Dim personName As String
    If folderType = "Sent" Then
        If mail.Recipients.Count > 0 Then
            personName = mail.Recipients.Item(1).Name
        Else
            personName = "NoRecipient"
        End If
    Else
        personName = mail.SenderName
        If Trim(personName) = "" Then personName = "NoSender"
    End If

    senderPart = CleanFileName(personName)
    If Len(senderPart) > 50 Then senderPart = Left(senderPart, 50)

    subjectPart = CleanFileName(mail.Subject)
    If Trim(subjectPart) = "" Then subjectPart = "NoSubject"

    Dim remainingLength As Long
    remainingLength = availableLength - Len(senderPart)

    If remainingLength < 20 Then
        senderPart = Left(senderPart, 30)
        remainingLength = availableLength - Len(senderPart)
        If remainingLength < 10 Then remainingLength = 10
    End If

    If Len(subjectPart) > remainingLength Then
        subjectPart = Left(subjectPart, remainingLength)
    End If

    Dim fileName As String
    fileName = dateTimePart & "_" & senderPart & "_" & subjectPart
    Do While Right(fileName, 1) = "_"
        fileName = Left(fileName, Len(fileName) - 1)
    Loop

    Dim fullPath As String
    fullPath = savePath & fileName & ".msg"

    ' 5) (중복 제거용) 이전 백업 경로를 "삭제하지 말고" 먼저 읽어만 둔다
    Dim oldPath As String
    oldPath = GetBackupFilePathFromMail(mail)

    ' Inbox 하위폴더로 이동되어 저장되는 케이스인지 판단
    Dim isMovedToSubfolder As Boolean
    isMovedToSubfolder = (folderType = "Inbox" And Trim(relativeFolderPath) <> "")

    ' 6) 저장
    mail.SaveAs fullPath, olMSG

    ' 7) 저장 검증 + 로깅 + (성공 시) oldPath 삭제 + (성공 시) BackupFilePath 갱신
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(fullPath) Then
        Dim fileSize As Long
        fileSize = fso.GetFile(fullPath).Size

        If fileSize < 100 Then
            LogError mail, "파일 크기 비정상: " & fullPath & " (" & fileSize & " bytes)"
        Else
            LogSuccess mail, fullPath, fileSize

            ' (핵심) 저장 성공한 경우에만, 이동 케이스라면 oldPath 삭제
            If isMovedToSubfolder Then
                If oldPath <> "" Then
                    ' 같은 경로면 삭제하지 않음
                    If StrComp(oldPath, fullPath, vbTextCompare) <> 0 Then
                        Dim fsoDel As Object
                        Set fsoDel = CreateObject("Scripting.FileSystemObject")

                        On Error Resume Next
                        If fsoDel.FileExists(oldPath) Then
                            fsoDel.DeleteFile oldPath, True
                        End If
                        On Error GoTo ErrorHandler

                        Set fsoDel = Nothing
                    End If
                End If
            End If

            ' 저장 성공한 경우에만 최신 백업 경로를 메일에 기록
            SetBackupFilePathToMail mail, fullPath
        End If
    Else
        LogError mail, "파일 저장 실패: " & fullPath
    End If

    Set fso = Nothing
    Exit Sub

ErrorHandler:
    LogError mail, "SaveMailAsMSG 에러: " & Err.Description
End Sub

'===========================================
' 메일 UserProperty에서 이전 백업 파일 경로 읽기
'===========================================
Private Function GetBackupFilePathFromMail(mail As Outlook.MailItem) As String
    On Error Resume Next
    Dim p As Outlook.UserProperty
    Set p = mail.UserProperties.Find(PROP_BACKUP_PATH, True)
    If Not p Is Nothing Then
        GetBackupFilePathFromMail = CStr(p.Value)
    Else
        GetBackupFilePathFromMail = ""
    End If
End Function

'===========================================
' 메일 UserProperty에 최신 백업 파일 경로 저장
'===========================================
Private Sub SetBackupFilePathToMail(mail As Outlook.MailItem, ByVal path As String)
    On Error Resume Next
    Dim p As Outlook.UserProperty
    Set p = mail.UserProperties.Find(PROP_BACKUP_PATH, True)
    If p Is Nothing Then
        Set p = mail.UserProperties.Add(PROP_BACKUP_PATH, olText, True)
    End If
    p.Value = path

    ' 필요 시 주석 처리 가능(환경에 따라 Save가 불필요하거나 정책에 막힐 수 있음)
    mail.Save
End Sub

'===========================================
' 폴더 경로 생성(재귀)
'===========================================
Public Sub CreateFolderPath(ByVal folderPath As String, Optional fso As Object = Nothing)
    Dim parentPath As String
    Dim needCleanup As Boolean

    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        needCleanup = True
    End If

    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    If fso.FolderExists(folderPath) Then
        If needCleanup Then Set fso = Nothing
        Exit Sub
    End If

    parentPath = fso.GetParentFolderName(folderPath)

    If parentPath <> "" Then
        If Not fso.FolderExists(parentPath) Then
            CreateFolderPath parentPath, fso
        End If
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    If needCleanup Then Set fso = Nothing
End Sub

'===========================================
' 파일명에 사용할 수 없는 문자 제거
'===========================================
Public Function CleanFileName(ByVal strFileName As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    Dim lastWasSpace As Boolean
    Dim lastWasUnderscore As Boolean

    result = ""
    lastWasSpace = False
    lastWasUnderscore = False

    For i = 1 To Len(strFileName)
        ch = Mid(strFileName, i, 1)

        Select Case ch
            Case "/", "\", ":", "*", "?", """", "<", ">", "|"
                If Not lastWasUnderscore Then
                    result = result & "_"
                    lastWasUnderscore = True
                    lastWasSpace = False
                End If

            Case vbCr, vbLf
                ' skip

            Case vbTab, " "
                If Not lastWasSpace Then
                    result = result & " "
                    lastWasSpace = True
                    lastWasUnderscore = False
                End If

            Case "_"
                If Not lastWasUnderscore Then
                    result = result & "_"
                    lastWasUnderscore = True
                    lastWasSpace = False
                End If

            Case Else
                result = result & ch
                lastWasSpace = False
                lastWasUnderscore = False
        End Select
    Next i

    CleanFileName = Trim(result)
End Function

'===========================================
' 폴더 경로("A\B\C") 정리: 세그먼트별로 CleanFileName 적용
'===========================================
Public Function CleanFolderPath(ByVal pathStr As String) As String
    Dim parts() As String
    Dim i As Long
    Dim cleaned As String
    Dim p As String

    parts = Split(pathStr, "\")

    cleaned = ""
    For i = LBound(parts) To UBound(parts)
        p = Trim(parts(i))
        If p <> "" Then
            p = CleanFileName(p)
            If p <> "" Then
                If cleaned = "" Then
                    cleaned = p
                Else
                    cleaned = cleaned & "\" & p
                End If
            End If
        End If
    Next i

    CleanFolderPath = cleaned
End Function

'===========================================
' 로그 파일 경로 (월별)
'===========================================
Public Function GetLogFilePath(Optional logType As String = "error") As String
    Dim logFolder As String
    logFolder = BACKUP_BASE_PATH & "logs\"
    CreateFolderPath logFolder
    GetLogFilePath = logFolder & Format(Now, "yyyy-mm") & "_" & logType & ".log"
End Function

'===========================================
' 에러 로깅
'===========================================
Public Sub LogError(mail As Outlook.MailItem, errMsg As String)
    On Error Resume Next

    Dim fso As Object
    Dim logFile As Object
    Dim logPath As String
    Dim logEntry As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = GetLogFilePath("error")
    Set logFile = fso.OpenTextFile(logPath, 8, True)

    logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
               errMsg & " | " & _
               SafeOneLine(mail.SenderName) & " | " & _
               SafeOneLine(mail.Subject)

    logFile.WriteLine logEntry
    logFile.Close

    Set logFile = Nothing
    Set fso = Nothing
End Sub

'===========================================
' 성공 로깅
'===========================================
Public Sub LogSuccess(mail As Outlook.MailItem, filePath As String, fileSize As Long)
    On Error Resume Next

    Dim fso As Object
    Dim logFile As Object
    Dim logPath As String
    Dim logEntry As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = GetLogFilePath("success")
    Set logFile = fso.OpenTextFile(logPath, 8, True)

    logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
               "SUCCESS" & " | " & _
               filePath & " | " & _
               CLng(fileSize) & " bytes" & " | " & _
               SafeOneLine(mail.SenderName) & " | " & _
               SafeOneLine(mail.Subject)

    logFile.WriteLine logEntry
    logFile.Close

    Set logFile = Nothing
    Set fso = Nothing
End Sub

Public Function SafeOneLine(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "|", "/")
    SafeOneLine = Trim(s)
End Function
