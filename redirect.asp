<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Query redirect page</title>
</head>

<body>


				<%
				dim keywords2, url, back
				
				keywords2=Request.Querystring("domesday-search_text")
				url= "/documentsonline/search-progress.asp?searchType=quicksearch&query="+ keywords2 +"&catID=24&pageNumber=1&queryType=1&mediaArray=*"
								
				if keywords <> "" Then
				response.redirect(url)
				Else
				response.redirect(url)				
				End if
				%>
				
</body>
</html>
