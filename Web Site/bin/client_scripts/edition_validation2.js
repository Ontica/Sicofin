<!--#INCLUDE FILE="../_scriptlibrary/rs.asp"-->
<script language="JavaScript" src="../programs/rs.htm"></script>
<script language="JavaScript">RSEnableRemoteScripting("../programs/")</script>
<script LANGUAGE="javascript">
<!--

function isDate(sDate) {
	var obj;
	obj = RSExecute("../programs/server_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../programs/server_scripts.asp","IsNumeric", sNumber, nDecimals);
	return obj.return_value;
}

//-->
</SCRIPT>
