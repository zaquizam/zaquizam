<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValBuscarHogarValidado.asp
	' 04ene21 - 05ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim idConsumo, valor
	'
	idConsumo = Request.Form("idConsumo")
	'
	Set rsscroll = CreateObject("ADODB.Recordset")
	Dim strSQLscroll, rsscroll, intRow
	strSQLscroll = "SELECT validado FROM PH_Consumo WHERE Id_Consumo =" & idConsumo
	rsscroll.open strSQLscroll, conexion
	valor = rsscroll("validado")
	'
	'Response.write rsscroll("validado")
	'response.end
	'
	If not rsscroll.EOF  Then
		Response.write rsscroll("validado")
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing
%>