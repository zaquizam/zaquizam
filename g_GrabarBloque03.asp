
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
	ytip=Request.QueryString("tip")
	yexp=Request.QueryString("exp")
	ymet=Request.QueryString("met")
	yamb=Request.QueryString("amb")
	yban=Request.QueryString("ban")
	yluz=Request.QueryString("luz")
	
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
	rsx3("Id_TipoVivienda") = ytip
	rsx3("OtroTipoVivienda") = yexp	
	rsx3("id_Metros") =	ymet	
	rsx3("NumeroAmbientes") = yamb	
	rsx3("NumeroBanos") = yban	
	rsx3("id_PuntosLuz") = yluz	
	
	rsx3.Update
	rsx3.Close 
	'set rsx3 = nothing 
	
	%>

	<%
	
%>