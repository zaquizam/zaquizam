<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvBuscarEstado.asp // 13ene21 - 
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'	
	' Buscar Los Hogares asociados al Estado
	'
	idHogar	=	Request.Querystring("id_hogar")
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " ss_Estado.Estado AS estado"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_PanelHogar"
	QrySql = QrySql & " INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado"
	QrySql = QrySql & " INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar = " & idHogar			
	'
	Set rsscroll = CreateObject("ADODB.Recordset")
	Dim strSQLscroll, rsscroll
	rsscroll.open QrySql, conexion
	'
	If not rsscroll.EOF  Then
		resultado = rsscroll("estado")
        response.write resultado
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing	
	'
%>