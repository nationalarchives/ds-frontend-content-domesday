<%

' //////////// SCRIPT /_russ_temp/domesday/email_friend.asp
' AUTHOR:											Russ Back (contractor)
'	DATE:												20 March 2006
' PURPOSE: 										Sends a text email via the TNA email component
' POST PARAMETERS:
' 	strToEmail	(string):			Email address to send to (i.e. "johnd@domain.com")
' 	strToName (string):				Name to send to (i.e. "John Doe")
'		strFromEmail (string):		Email address to send from (i.e. "janes@domain2.com")
'		strFromName (string):			Name to send from (i.e. "Jane Smith")
'		strScore (string):					Score formatted with commas if necessary (i.e. 6,566)
' RETURNS:
' 	boolSuccess (boolean):		True or False based on Error Count
'		strErrors (string):				Data validation or script error details

' EMAIL VALIDATION FUNCTION		emailValid(strEmailAddress)
' PURPOSE:										Validates email address against a regular expression
' SOURCE:											http://www.4guysfromrolla.com/webtech/090199-1.shtml
' PARAMETERS:									
'		strEmailAddress (string):	Email address (i.e. "johnd@domain.com")
' RETURNS:
'		emailValid (boolean):			True or False based on regular expression match

' ////////// END ////////////////////////////////////////////

Option Explicit
On Error Resume Next

Dim objMsg, strToEmail, strToName, strFromEmail, strFromName, strScore, boolSuccess, strErrors

' set/get variable values
boolSuccess = False
strToEmail = Trim(Request.Form("strToEmail"))
strToName = Trim(Request.Form("strToName"))
strFromEmail = Trim(Request.Form("strFromEmail"))
strFromName = Trim(Request.Form("strFromName"))
strScore = Trim(Request.Form("strScore"))

' validate data
If emailValid(strToEmail) = False Then strErrors = "Required form value strToEmail empty or invalid. "
If strToName = "" Then strErrors = strErrors & "Required form value strToName empty. "
If emailValid(strFromEmail) = False Then strErrors = strErrors & "Required form value strFromEmail empty or invalid. "
If strFromName = "" Then strErrors = strErrors & "Required form value strFromName empty. "
If strScore = "" Then strErrors = strErrors & "Required form value strScore empty. "

' process email if data ok
If strErrors = "" Then
	' Create the email message object
	Set objMsg = Server.CreateObject("NationalArchives.Email.Message")
	With objMsg
		' Set the message properties.
		.From = """" & strFromName & """ <" & strFromEmail & ">"
		.To = """" & strToName & """ <" & strToEmail & ">"
		.Subject = strFromName & " thought you'd like this..."
		.TextBody = "Hi " & strToName & "," & vbCrlf & vbCrlf &_
				strFromName & " has been playing a great Domesday game on The National Archives website and scored " & strScore &_
				". Think you can beat it? Go to http://www.nationalarchives.gov.uk/domesday/domesday-game/ to give it a go now."
		.Send
		' clean up resources in use by the message - always call this before nulling the object reference.
		.Close
	End With
	' destroy the object
	Set objMsg = Nothing
End If

' report script errors
If Err.Number <> 0 Then 
	strErrors = strErrors & "The script returned the following error: " & Err.Description & "."
End If
If strErrors = "" Then boolSuccess = True

' debug
' Response.Write(boolSuccess & ": " & strErrors)

' validate email address
Function emailValid(strEmailAddress)
  emailValid = False
  Dim objExpression, boolReturnVal
  Set objExpression = New RegExp
  ' Create regular expression:
  objExpression.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
  ' Set pattern:
  objExpression.IgnoreCase = true
  ' Set case sensitivity.
  boolReturnVal = objExpression.Test(strEmailAddress)
  ' Execute the search test.
  If Not boolReturnVal Then
    Exit Function
  End If
  emailValid = True
End Function

%>