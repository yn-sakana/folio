Attribute VB_Name = "FolioOutlook"
Option Explicit

' Create a reply draft in Outlook via COM.
' Only creates the draft - does NOT send.

Public Sub CreateReplyDraft(toAddress As String, subject As String, _
                            Optional bodyText As String = "", _
                            Optional selfAddress As String = "")
    Dim eh As New ErrorHandler: eh.Enter "FolioOutlook", "CreateReplyDraft"
    On Error GoTo ErrHandler

    Dim olApp As Object
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")

    Dim mail As Object
    Set mail = olApp.CreateItem(0) ' olMailItem

    mail.To = toAddress
    mail.Subject = subject
    If Len(bodyText) > 0 Then mail.body = bodyText

    ' Set sender if specified and multiple accounts exist
    If Len(selfAddress) > 0 Then
        Dim acct As Object
        For Each acct In olApp.Session.Accounts
            If LCase$(acct.SmtpAddress) = LCase$(selfAddress) Then
                Set mail.SendUsingAccount = acct
                Exit For
            End If
        Next acct
    End If

    mail.Display ' Show draft to user
    eh.OK: Exit Sub

ErrHandler: eh.Catch
End Sub
