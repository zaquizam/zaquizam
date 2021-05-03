
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>LLEGO25" 
	'response.end
	ynum=Request.QueryString("num") 
	yocu=Request.QueryString("ocu")
	yexp=Request.QueryString("exp")
	ymon=Request.QueryString("mon")
	
	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3

	sql = ""
	sql = sql & " Select * from PH_PanelHogar "
	sql = sql & " Where Id_PanelHogar = " & cint(ynum)
	'response.write "<br>220 sql:=" & sql
	'response.end
	rsx3.Open sql ,conexion
	rsx3("Id_OcupacionVivienda") = yocu
	rsx3("OtroOcupacionVivienda") =	yexp	
	rsx3("Id_MontoVivienda") =	ymon	
	rsx3.Update
	rsx3.Close 
	'set rsx3 = nothing 
	
%>