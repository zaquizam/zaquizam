
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
	ynumtv=Request.QueryString("numtv")
	ytiptv=Request.QueryString("tiptv")
	ysenal=Request.QueryString("senal")
	ycabl1=Request.QueryString("cabl1")
	ycabl2=Request.QueryString("cabl2")
	ytvon1=Request.QueryString("tvon1")
	ytvon2=Request.QueryString("tvon2")
	
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
	rsx3("Id_Televisores") = ynumtv
	rsx3("Id_TipoTelevisores") = ytiptv
	rsx3("Id_Senal") = ysenal
	rsx3("Id_Cablera1") = ycabl1
	rsx3("Id_Cablera2") = ycabl2
	rsx3("Id_TelevisionOnline1") = ytvon1
	rsx3("Id_TelevisionOnline2") = ytvon2
	rsx3.Update
	rsx3.Close 
	'set rsx3 = nothing 
	
%>