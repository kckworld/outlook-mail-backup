VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderItemsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1
Public folderType As String            ' "Inbox" / "Sent"
Public relativeFolderPath As String    ' 예: "프로젝트A\고객사1"

Private Sub Items_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item

        modMailBackup.SaveMailAsMSG mail, folderType, relativeFolderPath

        Set mail = Nothing
    End If

    Exit Sub

ErrorHandler:
    If TypeOf Item Is Outlook.MailItem Then
        modMailBackup.LogError Item, "Items_ItemAdd 에러(" & folderType & "\" & relativeFolderPath & "): " & Err.Description
    End If
End Sub

