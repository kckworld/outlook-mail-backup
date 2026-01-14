VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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


