Sub Document_Open()

    ' Declare variables
Dim olApp As Outlook.Application
Dim olNameSpace As NameSpace
Dim olAddressList As AddressList
Dim olAddressEntry As AddressEntry
Dim olAliasName As String
Dim olContact As Outlook.ExchangeUser

Dim Name As String
Dim Email As String
Dim PhoneNumber As String

    ' Set up Outlook
Set olApp = CreateObject("Outlook.Application")
Set olNameSpace = olApp.GetNamespace("MAPI")
Set olAddressList = olNameSpace.AddressLists("Zealand Business College")

Set olAddressEntry = olAddressList.AddressEntries(olNameSpace.CurrentUser.Name)

If olAddressEntry Is Nothing Then
    MsgBox ("Kunne ikke finde bruger i " & olAddressList)
End If

Set olContact = olAddressEntry.GetExchangeUser()

If olContact Is Nothing Then
    MsgBox ("Kunne ikke hente brugerens oplysninger")
Else
    ' Get relevant information
    Name = olContact.FirstName & " " & olContact.LastName
    Email = olContact.PrimarySmtpAddress
    PhoneNumber = olContact.MobileTelephoneNumber
End If

    ' Navigate to bookmark and edit text
    
If ActiveDocument.Bookmarks.Exists("bm_dato") = True Then
    ActiveDocument.Bookmarks("bm_dato").Select
    Selection.TypeText (Date)
    ActiveDocument.Bookmarks("bm_dato").Delete
End If

If ActiveDocument.Bookmarks.Exists("bm_navn") = True Then
    ActiveDocument.Bookmarks("bm_navn").Select
    Selection.TypeText (Name)
End If

If ActiveDocument.Bookmarks.Exists("bm_email") = True Then
    ActiveDocument.Bookmarks("bm_email").Select
    Selection.TypeText (Email)
End If

If ActiveDocument.Bookmarks.Exists("bm_tlf") = True Then
    ActiveDocument.Bookmarks("bm_tlf").Select
    Selection.TypeBackspace
    Selection.TypeText (PhoneNumber)
End If

End Sub
