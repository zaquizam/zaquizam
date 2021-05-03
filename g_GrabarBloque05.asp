
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
	yagub=Request.QueryString("agub")
	yagun=Request.QueryString("agun")
	yase=Request.QueryString("ase")
	yele=Request.QueryString("ele")
	ytel=Request.QueryString("tel")
	
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
	rsx3("Id_AguasBlancas") = yagub
	rsx3("Id_AguasNegras") = yagun
	rsx3("Id_AseoUrbano") = yase
	rsx3("Id_ServicioElectricidad") = yele
	rsx3("Id_ServicioTelefono") = ytel
	rsx3.Update
	rsx3.Close 
	'set rsx3 = nothing 
	
%>