<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Query redirect page</title>
</head>

<body>
<!-- Google Tag Manager -->
<noscript><iframe src="//www.googletagmanager.com/ns.html?id=GTM-T8DSWV"
				  height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
		new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
		j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
		'//www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
	})(window,document,'script','dataLayer','GTM-T8DSWV');</script>
<!-- End Google Tag Manager -->



				<%
				dim keywords, url, back
				
				keywords=Request.Querystring("keywords")
				url= "/documentsonline/search-progress.asp?searchType=browserefine&query=scope%3d" & keywords & "&catID=24&pageNumber=1&queryType=1&mediaArray=*"
				
				if keywords <> "" Then
				response.redirect(url)
				Else
				response.redirect(url)				
				End if
				%>
				
</body>
</html>
