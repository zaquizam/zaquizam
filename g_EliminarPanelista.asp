<%@language=vbscript%>

<!--#include file="Conexion.asp"-->
 
<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================

	sBus=Request.QueryString("num")

	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3

	sql = ""
    sql = sql & " Delete "
	sql = sql & " FROM PH_Panelistas"
	sql = sql & " WHERE "
	sql = sql & " Id_Panelista = " & sBus
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx3.Open sql ,conexion
	'response.end
		%>

	
	<%
%>
