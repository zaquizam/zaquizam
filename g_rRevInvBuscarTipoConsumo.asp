<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvBuscarTipoConsumo.asp // 13ene21 - 
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'	
	' Buscar el Tipo de Consumo
	'
	idConsumo =	Request.Querystring("id_consumo")
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_TipoConsumo.TipoConsumo AS tipo"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN cacevedo_atenas.PH_TipoConsumo ON cacevedo_atenas.PH_Consumo.id_TipoConsumo = cacevedo_atenas.PH_TipoConsumo.Id_TipoConsumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " cacevedo_atenas.PH_Consumo.Id_Consumo = " & idConsumo

	'
	Set rsscroll = CreateObject("ADODB.Recordset")
	Dim strSQLscroll, rsscroll
	rsscroll.open QrySql, conexion
	'
	If not rsscroll.EOF  Then
		resultado = rsscroll("tipo")
        response.write resultado
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing	
	'
%>