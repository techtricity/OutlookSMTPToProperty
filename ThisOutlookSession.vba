Private WithEvents m_Inbox As Outlook.Items
Private WithEvents m_SentItems As Outlook.Items

Public Sub Application_Startup()
  Dim Session As Outlook.NameSpace
  Set Session = Application.Session

  Set m_Inbox = Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub m_Inbox_ItemAdd(ByVal Item As Object)
  Dim UserProps As Outlook.UserProperties
  Dim Prop As Outlook.UserProperty
  Dim FieldName As String
  
  Dim Header As String
  Dim a() As String
  Dim ToRecipient As String
  Dim pa As Outlook.PropertyAccessor

  If TypeOf Item Is Outlook.MailItem Then
    Header = GetInetHeaders(Item)

    a = Split(Header, vbCrLf)
    For i = 0 To UBound(a)
      If InStr(1, a(i), "To:", vbTextCompare) = 1 Then
        If InStr(1, a(i), "<", vbTextCompare) > 0 Then
          ToRecipient = Replace(Split(a(i), "<")(1), ">", "")
        Else
          ToRecipient = Trim(Split(a(i), "To:")(1))
        End If
      End If
    Next
  
    FieldName = "SMTP To"
    Set UserProps = Item.UserProperties
    Set Prop = UserProps.Find(FieldName, True)
    If Prop Is Nothing Then
      Set Prop = UserProps.Add(FieldName, olText, True)
    End If
    Prop.Value = ToRecipient
    Item.Save
  End If
End Sub

Function GetInetHeaders(olkMsg As Outlook.MailItem) As String
    ' Purpose: Returns the internet headers of a message.'
    ' Written: 4/28/2009'
    ' Author:  BlueDevilFan'
    ' //techniclee.wordpress.com/
    ' Outlook: 2007'
    Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Dim olkPA As Outlook.PropertyAccessor
    Set olkPA = olkMsg.PropertyAccessor
    GetInetHeaders = olkPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    Set olkPA = Nothing
End Function
