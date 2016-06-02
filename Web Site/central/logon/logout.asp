<%
	Option Explicit  
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Session.Abandon()	
	
%>
<html>
<head>
<script LANGUAGE=javascript>
<!--
   function window_onload() {
		window.open("/empiria/default.asp", "_top");
   }

//-->
</SCRIPT>
<body LANGUAGE=javascript onload="return window_onload()">
</body>
</html>
