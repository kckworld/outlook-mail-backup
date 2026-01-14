Option Explicit

Private FolderHandlers As Collection
Private InboxRoot As Outlook.folder
Private SentRoot As Outlook.folder

Private Sub Application_Startup()
    InitializeEventHandler
End Sub

Public Sub InitializeEventHandler()
    On Error GoTo ErrorHandler

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Set FolderHandlers = New Collection

    Set InboxRoot = ns.GetDefaultFolder(olFolderInbox)
    Set SentRoot = ns.GetDefaultFolder(olFolderSentMail)

    ' Inbox: 루트 + 하위 전체 감시 (루트도 백업됨. 단, 하위로 이동 시 루트 백업파일 자동 삭제됨)
    RegisterFolderAndSubfolders InboxRoot, "Inbox", InboxRoot, False

    ' Sent: 루트만 감시 (원하면 False로 바꾸면 하위까지 감시)
    RegisterFolderAndSubfolders SentRoot, "Sent", SentRoot, True

    Set ns = Nothing

    MsgBox "자동 백업 이벤트 활성화 완료" & vbCrLf & _
           "Inbox(하위 포함), Sent(보낸편지함) 모니터링 중" & vbCrLf & _
           "저장 루트: " & modMailBackup.BACKUP_BASE_PATH, vbInformation, "Outlook Mail Backup"
    Exit Sub

ErrorHandler:
    MsgBox "InitializeEventHandler 오류: " & Err.Description, vbCritical, "Outlook Mail Backup"
End Sub

Private Sub RegisterFolderAndSubfolders( _
    ByVal targetFolder As Outlook.folder, _
    ByVal folderType As String, _
    ByVal rootFolder As Outlook.folder, _
    ByVal rootOnly As Boolean)

    On Error Resume Next

    Dim h As clsFolderItemsHandler
    Set h = New clsFolderItemsHandler

    Set h.Items = targetFolder.Items
    h.folderType = folderType
    h.relativeFolderPath = GetRelativePathUnderRoot(targetFolder, rootFolder)

    FolderHandlers.Add h

    If rootOnly Then Exit Sub

    Dim subFolder As Outlook.folder
    For Each subFolder In targetFolder.Folders
        RegisterFolderAndSubfolders subFolder, folderType, rootFolder, rootOnly
    Next subFolder
End Sub

Private Function GetRelativePathUnderRoot(ByVal folder As Outlook.folder, ByVal rootFolder As Outlook.folder) As String
    On Error GoTo SafeExit

    If folder.EntryID = rootFolder.EntryID Then
        GetRelativePathUnderRoot = ""
        Exit Function
    End If

    Dim rel As String
    rel = ""

    Dim cur As Outlook.folder
    Set cur = folder

    Do While Not (cur Is Nothing)
        If cur.EntryID = rootFolder.EntryID Then Exit Do

        If rel = "" Then
            rel = cur.Name
        Else
            rel = cur.Name & "\" & rel
        End If

        Set cur = cur.Parent
    Loop

    GetRelativePathUnderRoot = rel
    Exit Function

SafeExit:
    GetRelativePathUnderRoot = ""
End Function


'===========================================
' 수동 실행: 선택한 메일 저장
' 사용법: Outlook에서 메일 여러 개 선택 -> Alt+F8 -> SaveSelectedMailsAsMSG 실행
'===========================================
Public Sub SaveSelectedMailsAsMSG()
    On Error GoTo ErrorHandler

    ' 루트 폴더 핸들이 아직 없으면 초기화(Outlook 재시작 직후 등)
    If InboxRoot Is Nothing Or SentRoot Is Nothing Then
        Dim ns As Outlook.NameSpace
        Set ns = Application.GetNamespace("MAPI")
        Set InboxRoot = ns.GetDefaultFolder(olFolderInbox)
        Set SentRoot = ns.GetDefaultFolder(olFolderSentMail)
        Set ns = Nothing
    End If

    Dim selectedItems As Outlook.Selection
    Set selectedItems = Application.ActiveExplorer.Selection

    If selectedItems Is Nothing Or selectedItems.Count = 0 Then
        MsgBox "저장할 메일을 선택해주세요.", vbExclamation, "Outlook Mail Backup"
        Exit Sub
    End If

    Dim totalCount As Long
    Dim savedCount As Long
    totalCount = selectedItems.Count
    savedCount = 0

    If totalCount > 50 Then
        If MsgBox(totalCount & "개의 메일을 백업합니다. 진행하시겠습니까?", vbYesNo + vbQuestion, "Outlook Mail Backup") = vbNo Then
            Exit Sub
        End If
    End If

    Dim it As Object
    For Each it In selectedItems
        If TypeOf it Is Outlook.MailItem Then
            Dim mail As Outlook.MailItem
            Set mail = it

            Dim parentFolder As Outlook.folder
            Set parentFolder = mail.Parent

            Dim folderType As String
            Dim relPath As String
            folderType = "Inbox"
            relPath = ""

            ' SentRoot 하위면 Sent로 처리
            If Not SentRoot Is Nothing Then
                If IsFolderUnderRoot(parentFolder, SentRoot) Then
                    folderType = "Sent"
                    relPath = GetRelativePathUnderRoot(parentFolder, SentRoot)
                ElseIf Not InboxRoot Is Nothing Then
                    ' InboxRoot 하위면 Inbox로 처리
                    If IsFolderUnderRoot(parentFolder, InboxRoot) Then
                        folderType = "Inbox"
                        relPath = GetRelativePathUnderRoot(parentFolder, InboxRoot)
                    Else
                        ' 그 외 폴더는 기본 Inbox 처리(원하면 여기서 분기 변경 가능)
                        folderType = "Inbox"
                        relPath = ""
                    End If
                End If
            End If

            modMailBackup.SaveMailAsMSG mail, folderType, relPath
            savedCount = savedCount + 1

            Set mail = Nothing
            Set parentFolder = Nothing

            DoEvents
        End If
    Next it

    MsgBox savedCount & "개의 메일이 저장되었습니다." & vbCrLf & _
           "경로: " & modMailBackup.BACKUP_BASE_PATH, vbInformation, "Outlook Mail Backup"
    Exit Sub

ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical, "Outlook Mail Backup"
End Sub

'===========================================
' folder가 rootFolder 아래(또는 동일)인지 판별
'===========================================
Private Function IsFolderUnderRoot(ByVal folder As Outlook.folder, ByVal rootFolder As Outlook.folder) As Boolean
    On Error GoTo SafeExit

    Dim cur As Outlook.folder
    Set cur = folder

    Do While Not (cur Is Nothing)
        If cur.EntryID = rootFolder.EntryID Then
            IsFolderUnderRoot = True
            Exit Function
        End If
        Set cur = cur.Parent
    Loop

SafeExit:
    IsFolderUnderRoot = False
End Function


