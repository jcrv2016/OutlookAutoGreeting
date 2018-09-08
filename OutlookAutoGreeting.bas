Attribute VB_Name = "Module1"
Sub OutlookAutoGreeting()

Dim origEmail As MailItem: Set origEmail = ActiveExplorer.Selection(1)
Dim replyEmail As MailItem: Set replyEmail = origEmail.ReplyAll

'Get current time
Dim LHour As Integer: LHour = Hour(Now)

'Pull SenderName from origEmail
Dim SenderName As String: SenderName = Split(origEmail.SenderName)(0)

'Ignore SenderName original case, make first char uppercase, all others lowercase
SenderName1 = LCase(Right(SenderName, (Len(SenderName) - 1)))
SenderName2 = UCase(Left(SenderName, 1))
SenderName = SenderName2 & SenderName1

'Generate time-dependent salutation
Dim Morning: Morning = "Good morning "
Dim Afternoon: Afternoon = "Good afternoon "
Dim Evening: Evening = "Good evening "
Dim TimeOfDay As String

If (LHour <= 11) Then
    TimeOfDay = Morning
ElseIf (LHour <= 4) Then
    TimeOfDay = Afternoon
Else
    TimeOfDay = Evening
End If

'Append salutation with name
Dim Greeting As String: Greeting = TimeOfDay & SenderName & ","

'Assemble/display email content
replyEmail.HTMLBody = Greeting & vbNewLine & replyEmail.HTMLBody & origEmail.Reply.HTMLBody
replyEmail.Display

End Sub
