<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_rRevInvBuscarDiaSemana.asp // 14ene21 -
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'	
	' Buscar la Fecha y dia del Consumo
	'
	idConsumo =	Request.Querystring("id_consumo")
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " (CASE DATENAME(dw,fecha_creacion) when 'Monday' then 'LUN' when 'Tuesday' then 'MAR' when 'Wednesday' then 'MIE' when 'Thursday' then 'JUE' when 'Friday' then 'VIE' when 'Saturday' then 'SAB' when 'Sunday' then 'DOM' END) AS DIA,"
	QrySql = QrySql & " FORMAT (PH_Consumo.fecha_creacion, 'dd-MM-yyyy ') AS FECHA"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.id_Consumo = " & idConsumo
	'
	Set rsscroll = CreateObject("ADODB.Recordset")
	Dim strSQLscroll, rsscroll
	rsscroll.open QrySql, conexion
	'
	If not rsscroll.EOF  Then		
		resultado= CStr(rsscroll(0)) & " - " & CStr(rsscroll(1))
        response.write resultado
	Else
		Response.write False
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing	
	'
	''
%>