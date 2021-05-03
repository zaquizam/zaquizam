<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvBuscarSemanas.asp // 13ene21 - 
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsSemanas, arrSemanas
	'	
	' Buscar Los Hogares asociados al Estado
	'
	idConsumo	=	Request.Querystring("id_consumo")
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " ss_Semana.Semana as semana"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Consumo = " & idConsumo
	'
	Set rsscroll = CreateObject("ADODB.Recordset")
	Dim strSQLscroll, rsscroll
	rsscroll.open QrySql, conexion
	'
	If not rsscroll.EOF  Then
		resultado = rsscroll("semana")
        response.write resultado
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing	
	'
%>