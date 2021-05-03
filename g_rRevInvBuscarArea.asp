<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvBuscarArea.asp // 13ene21 - 
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
	QrySql = QrySql & " PH_GArea.Area AS area"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_PanelHogar"
	QrySql = QrySql & " INNER JOIN cacevedo_atenas.PH_GAreaEstado ON cacevedo_atenas.PH_PanelHogar.Id_Estado = cacevedo_atenas.PH_GAreaEstado.Id_Estado"
	QrySql = QrySql & " INNER JOIN cacevedo_atenas.PH_GArea ON cacevedo_atenas.PH_GAreaEstado.Id_Area = cacevedo_atenas.PH_GArea.Id_Area"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " cacevedo_atenas.PH_PanelHogar.Id_PanelHogar = " & idHogar	
	'
	Set rsscroll = CreateObject("ADODB.Recordset")
	Dim strSQLscroll, rsscroll
	rsscroll.open QrySql, conexion
	'
	If not rsscroll.EOF  Then
		resultado = rsscroll("area")
        response.write resultado
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing	
	'
%>