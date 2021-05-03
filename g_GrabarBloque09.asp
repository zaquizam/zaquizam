
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
	ymas=Request.QueryString("mas")
	yper=Request.QueryString("per")
	ygat=Request.QueryString("gat")
	ypez=Request.QueryString("pez")
	yave=Request.QueryString("ave")
	yroe=Request.QueryString("roe")
	yotr=Request.QueryString("otr")
	
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
	rsx3("Id_Mascotas") = ymas
	rsx3("Ind_Perro") = yper
	rsx3("Ind_Gato") = ygat
	rsx3("Ind_Pez") = ypez
	rsx3("Ind_Ave") = yave
	rsx3("Ind_Roedor") = yroe
	rsx3("Ind_Otro") = yotr
	rsx3.Update
	rsx3.Close 
	'set rsx3 = nothing 
	
%>