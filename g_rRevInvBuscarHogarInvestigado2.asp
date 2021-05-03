<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValBuscarHogarInvestigado.asp
	' 10ene21 - 
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim idConsumo, valor
	'
	idConsumo = Request.QueryString("idConsumo")
	'
	Dim strSQLscroll, rsscroll
	Set rsscroll = CreateObject("ADODB.Recordset")	
	strSQLscroll = "SELECT enviado_investigar FROM PH_Consumo WHERE enviado_investigar=1 AND Respuesta_investigacion=0 and Id_Consumo =" & idConsumo
	rsscroll.open strSQLscroll, conexion
	'
	' Response.write strSQLscroll
	' Response.end
	'
	If not rsscroll.EOF  Then
		Response.write rsscroll("enviado_investigar")
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing
	'
%>