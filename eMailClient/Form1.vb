Imports MailKit
Imports MimeKit
Imports MailKit.Net.Smtp
Imports MailKit.Net.Imap
Imports MailKit.Search
Imports System.IO

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Send email with plain text body
        Dim message = New MimeMessage()
        message.From.Add(New MailboxAddress("xxxx", "xxxx"))
        message.[To].Add(New MailboxAddress("xxxx", "xxxx"))
        message.Subject = "xxxx"
        message.Body = New TextPart("plain") With {
            .Text = "xxxx"}
        Using client = New SmtpClient()
            client.Connect("smtp.office365.com", 587, False)
            client.Authenticate("xxxx", "xxxx")
            client.Send(message)
            client.Disconnect(True)
        End Using

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Read email and filter the list based on subject line
        Using client = New ImapClient()
            client.Connect("outlook.office365.com", 993, True)
            client.Authenticate("xxxx", "xxxx")
            Dim inbox = client.Inbox
            inbox.Open(FolderAccess.[ReadOnly])
            Label1.Text = "Total messages: " & inbox.Count
            Label2.Text = "Recent messages: " & inbox.Recent

            For i As Integer = 0 To inbox.Count - 1
                Dim message = inbox.GetMessage(i)
                ListBox1.Items.Add("Subject: " & message.Subject & "----------" & message.Date.ToString)
                If message.Subject = "xxxx" Then
                    MessageBox.Show(message.TextBody)
                End If
            Next
            client.Disconnect(True)
        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Changing the Read Flag of latest email
        Using client = New ImapClient()
            client.Connect("outlook.office365.com", 993, True)
            client.Authenticate("xxxx", "xxxx")
            Dim inbox = client.Inbox
            inbox.Open(FolderAccess.[ReadWrite])
            inbox.AddFlags(inbox.Count - 1, MessageFlags.Seen, True)
            client.Disconnect(True)
        End Using
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Using client = New ImapClient()
            client.Connect("outlook.office365.com", 993, True)
            client.Authenticate("xxxx", "xxxx")
            Dim inbox = client.Inbox
            inbox.Open(FolderAccess.[ReadWrite])
            Label1.Text = "Total messages: " & inbox.Count
            Label2.Text = "Recent messages: " & inbox.Recent

            ' search for messages where the Subject header contains either "WARNING"
            Dim query = SearchQuery.SubjectContains("Settings")
            Dim uids = client.Inbox.Search(query)

            ' fetch summary information for the search results (we will want the UID and the BODYSTRUCTURE
            ' of each message so that we can extract the text body and the attachments)
            Dim items = client.Inbox.Fetch(uids, MessageSummaryItems.Flags)

            For Each item In items
                ListBox1.Items.Add("Message # " & item.Index & " has flags: " & item.Flags.Value)

                If item.Flags.Value.HasFlag(MessageFlags.Seen) Then
                    ListBox1.Items.Add("The message has been read.")
                    inbox.AddFlags(item.UniqueId, MessageFlags.Deleted, True)
                Else
                    ListBox1.Items.Add("The message has not been read.")
                    inbox.AddFlags(item.UniqueId, MessageFlags.Seen, True)
                End If
            Next
            client.Disconnect(True)
        End Using
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim message = New MimeMessage()
        message.From.Add(New MailboxAddress("xxxx", "xxxx"))
        message.[To].Add(New MailboxAddress("xxxx", "xxxx"))
        message.Subject = "xxxx"

        Dim eMailBody As New BodyBuilder
        eMailBody.TextBody = "Please see attached."
        eMailBody.Attachments.Add("xxxx")
        message.Body = eMailBody.ToMessageBody

        Using client = New SmtpClient()
            client.Connect("smtp.office365.com", 587, False)
            client.Authenticate("xxxx", "xxxx")
            client.Send(message)
            client.Disconnect(True)
        End Using
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Using client = New ImapClient()
            client.Connect("outlook.office365.com", 993, True)
            client.Authenticate("xxxx", "xxxx")
            Dim inbox = client.Inbox
            inbox.Open(FolderAccess.ReadWrite)
            Label1.Text = "Total messages: " & inbox.Count
            Label2.Text = "Recent messages: " & inbox.Recent

            For i As Integer = inbox.Count - 1 To 0 Step -1
                Dim message = inbox.GetMessage(i)
                If message.Subject = "xxxx" Then
                    For Each attachment As MimeEntity In message.Attachments
                        Using stream = File.Create("xxxx")
                            If TypeOf attachment Is MessagePart Then
                                Dim rfc822 = CType(attachment, MessagePart)
                                rfc822.Message.WriteTo(stream)
                            Else
                                Dim part = CType(attachment, MimePart)
                                part.Content.DecodeTo(stream)
                            End If
                        End Using
                    Next
                    inbox.AddFlags(i, MessageFlags.Deleted, True)
                    Exit For
                End If
            Next
            client.Disconnect(True)
        End Using
    End Sub

End Class